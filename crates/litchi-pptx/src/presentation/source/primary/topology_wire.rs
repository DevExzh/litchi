//! Bounded durable transport for source-backed primary topology changes.
//!
//! Version 3 carries both complete source states.  A state includes the
//! presentation root, its stable catalog order, the topology proof supplied by
//! the owner, and the exact root metadata images.  Resource changes carry the
//! exact before/after images for one stable slide identity.  Version 5
//! (record 0303) appends the Parts that removed slides owned exclusively:
//! notes slides, pictures, diagrams, VML drawings, comments and embedded
//! objects, each with its exact payload and relationship images.  A patch
//! without such Parts is still encoded as version 3.  This module only
//! checks transport integrity and bounds; semantic XML, relationship and
//! ownership validation belongs to the application layer.

use super::super::MAX_SOURCE_BACKED_SLIDE_BATCH;
use crate::presentation::order::{StableSlideKey, stable_slide_key};
use crate::{Error, Result};
use litchi_opc::{PackURI, SourceLimitKind, SourceLimits};
use sha2::{Digest, Sha256};
use std::collections::HashSet;
use std::sync::Arc;

const MAGIC: &[u8; 8] = b"LPSX0003";
const VERSION: u32 = 3;
/// Version 5 adds owned dependent Parts. `LPSX0004` is the durable envelope
/// for shape removals, so the topology wire skips that number.
pub(super) const DEPENDENT_MAGIC: &[u8; 8] = b"LPSX0005";
const DEPENDENT_VERSION: u32 = 5;
const CHECKSUM_BYTES: usize = 32;
const CHECKSUM_DOMAIN: &[u8] = b"litchi-pptx-source-topology-wire-v3\0";
const DEPENDENT_CHECKSUM_DOMAIN: &[u8] = b"litchi-pptx-source-topology-wire-v5\0";
/// The most dependent Parts one patch may carry.
pub(super) const MAX_DEPENDENTS: usize = 4096;

// magic + version + three counts + two state headers + checksum.  A state
// header contains three payload lengths and one topology digest.
const HEADER_BYTES: usize = 8 + 4 + (4 * 3) + ((8 * 3 + 32) * 2);
const KEY_FIXED_BYTES: usize = 4 + 4 + 4;
const CHANGE_FIXED_BYTES: usize = KEY_FIXED_BYTES + 4 + 4 + 1 + 1;
const MAX_WIRE_BYTES: usize = 128 * 1024 * 1024;

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct Data {
    pub(super) before: State,
    pub(super) after: State,
    pub(super) changes: Arc<[Change]>,
    /// Parts owned exclusively by removed or restored slides, sorted by
    /// folded Part name.
    pub(super) dependents: Arc<[Dependent]>,
}

/// The exact before/after image of one Part that a removed slide owned
/// exclusively (record 0303). `before` is present when the Part exists in
/// the state before the patch, `after` when it exists after it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct Dependent {
    pub(super) part_uri: PackURI,
    pub(super) content_type: String,
    pub(super) before: Option<PartImage>,
    pub(super) after: Option<PartImage>,
}

impl Dependent {
    pub(super) fn inverse(&self) -> Self {
        Self {
            part_uri: self.part_uri.clone(),
            content_type: self.content_type.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct State {
    pub(super) root: Arc<Vec<u8>>,
    pub(super) order: Arc<[StableSlideKey]>,
    pub(super) topology_digest: [u8; 32],
    pub(super) presentation_relationships: Arc<Vec<u8>>,
    pub(super) content_types: Arc<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct Change {
    pub(super) key: StableSlideKey,
    pub(super) content_type: String,
    pub(super) relationship_type: String,
    pub(super) before: Option<PartImage>,
    pub(super) after: Option<PartImage>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct PartImage {
    pub(super) xml: Arc<Vec<u8>>,
    pub(super) relationships: Option<Arc<Vec<u8>>>,
}

/// Lengths used by a topology replacement preflight.  The preflight must be
/// able to account for a source image before retaining its bytes; keeping
/// only lengths here avoids constructing a second owned image merely to
/// discover that the durable wire ceiling cannot admit it.
#[derive(Clone, Copy)]
pub(super) struct ImageLengths {
    pub(super) xml: usize,
    pub(super) relationships: Option<usize>,
}

#[derive(Clone)]
pub(super) struct ChangeLengths {
    pub(super) key: StableSlideKey,
    pub(super) content_type: String,
    pub(super) relationship_type: String,
    pub(super) before: Option<ImageLengths>,
    pub(super) after: Option<ImageLengths>,
}

#[derive(Clone)]
pub(super) struct DependentLengths {
    pub(super) part_uri: PackURI,
    pub(super) content_type: String,
    pub(super) before: Option<ImageLengths>,
    pub(super) after: Option<ImageLengths>,
}

#[derive(Clone, Copy)]
pub(super) struct StateLengths {
    pub(super) root: usize,
    pub(super) presentation_relationships: usize,
    pub(super) content_types: usize,
}

impl Data {
    pub(super) fn is_noop(&self) -> bool {
        self.before == self.after && self.changes.is_empty() && self.dependents.is_empty()
    }

    /// Reverse complete topology and resource state while retaining every
    /// immutable byte image through shared ownership.
    pub(super) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            changes: self
                .changes
                .iter()
                .map(|change| Change {
                    key: change.key.clone(),
                    content_type: change.content_type.clone(),
                    relationship_type: change.relationship_type.clone(),
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
            dependents: self.dependents.iter().map(Dependent::inverse).collect(),
        }
    }

    /// Encode a bounded version-3 source topology patch.
    pub(super) fn to_bytes(&self) -> Result<Vec<u8>> {
        let length = self.encoded_len()?;
        let mut output = Vec::new();
        output.try_reserve_exact(length).map_err(allocation)?;
        write_header(&mut output, self)?;
        write_order(&mut output, &self.before.order)?;
        write_order(&mut output, &self.after.order)?;
        write_state_payload(&mut output, &self.before);
        write_state_payload(&mut output, &self.after);
        for change in self.changes.iter() {
            write_key(&mut output, &change.key)?;
            write_string(&mut output, &change.content_type)?;
            write_string(&mut output, &change.relationship_type)?;
            write_image(&mut output, change.before.as_ref())?;
            write_image(&mut output, change.after.as_ref())?;
        }
        if !self.dependents.is_empty() {
            put_u32(&mut output, self.dependents.len())?;
            for dependent in self.dependents.iter() {
                write_string(&mut output, dependent.part_uri.as_str())?;
                write_string(&mut output, &dependent.content_type)?;
                write_image(&mut output, dependent.before.as_ref())?;
                write_image(&mut output, dependent.after.as_ref())?;
            }
        }
        let checksum = checksum(&output, self.version());
        output.extend_from_slice(&checksum);
        if output.len() != length {
            return Err(invalid("wire length accounting mismatch"));
        }
        Ok(output)
    }

    /// Decode a bounded version-3 source topology patch under an explicit
    /// source policy.  The first pass checks every count, length, UTF-8
    /// string, marker, aggregate and framing boundary before the retained
    /// vectors and payload images are allocated.  The second pass materializes
    /// only after that preflight succeeds.
    pub(super) fn from_bytes(bytes: &[u8], limits: SourceLimits) -> Result<Self> {
        if bytes.len() > MAX_WIRE_BYTES {
            return Err(wire_limit(bytes.len()));
        }
        let body_len = bytes
            .len()
            .checked_sub(CHECKSUM_BYTES)
            .ok_or_else(|| invalid("truncated checksum"))?;
        if body_len < HEADER_BYTES {
            return Err(invalid("truncated header"));
        }
        let expected: [u8; 32] = bytes[body_len..]
            .try_into()
            .map_err(|_| invalid("checksum width"))?;
        let version = match bytes.get(..MAGIC.len()) {
            Some(magic) if magic == MAGIC => Version::Three,
            Some(magic) if magic == DEPENDENT_MAGIC => Version::Five,
            _ => return Err(invalid("unsupported wire version")),
        };
        if checksum(&bytes[..body_len], version) != expected {
            return Err(invalid("checksum mismatch"));
        }

        let body = &bytes[..body_len];
        let mut scan = Input {
            bytes: body,
            offset: 0,
        };
        let header = read_header(&mut scan, limits, version)?;
        preflight_body(&mut scan, &header, limits, version)?;
        if scan.offset != body.len() {
            return Err(invalid("trailing bytes"));
        }

        let mut input = Input {
            bytes: body,
            offset: 0,
        };
        let header = read_header(&mut input, limits, version)?;
        let before_order = read_order(&mut input, header.before_count, limits)?;
        let after_order = read_order(&mut input, header.after_count, limits)?;
        let before = read_state(&mut input, header.before, before_order, limits)?;
        let after = read_state(&mut input, header.after, after_order, limits)?;
        let changes = read_changes(&mut input, header.change_count, limits)?;
        let dependents = match version {
            Version::Three => Vec::new(),
            Version::Five => read_dependents(&mut input, limits)?,
        };
        if input.offset != body.len() {
            return Err(invalid("trailing bytes"));
        }
        if version == Version::Five && dependents.is_empty() {
            return Err(invalid("version 5 carries no dependent Part"));
        }
        let data = Self {
            before,
            after,
            changes: changes.into(),
            dependents: dependents.into(),
        };
        validate_data(&data, limits)?;
        Ok(data)
    }

    /// Compute the exact encoded size, including the SHA-256 checksum.
    pub(super) fn encoded_len(&self) -> Result<usize> {
        // Encoding has no caller-supplied policy.  The durable wire still
        // admits only values accepted by the finite default source profile;
        // callers decoding under a tighter profile get that profile in
        // `from_bytes`.
        validate_data(self, SourceLimits::default())?;
        Self::encoded_len_from_data(self)
    }

    pub(super) fn encoded_len_from_data(data: &Data) -> Result<usize> {
        let changes = data
            .changes
            .iter()
            .map(change_lengths)
            .collect::<Result<Vec<_>>>()?;
        let dependents = data
            .dependents
            .iter()
            .map(dependent_lengths)
            .collect::<Result<Vec<_>>>()?;
        encoded_len_from_lengths(
            &data.before.order,
            &data.after.order,
            state_lengths(&data.before),
            state_lengths(&data.after),
            &changes,
            &dependents,
        )
    }

    /// Compute the exact wire length after replacing one slide change.  The
    /// replacement is evaluated from lengths only, so callers can preflight a
    /// source image before `physical_image` copies its payload into retained
    /// patch state.  Existing changes retain their metadata and all unrelated
    /// changes/dependents are accounted for exactly as in `encoded_len`.
    pub(super) fn encoded_len_with_replacement(
        &self,
        key: &StableSlideKey,
        before: Option<ImageLengths>,
        after: Option<ImageLengths>,
        content_type: &str,
        relationship_type: &str,
    ) -> Result<usize> {
        let mut changes = self
            .changes
            .iter()
            .map(change_lengths)
            .collect::<Result<Vec<_>>>()?;
        let mut replaced = false;
        for change in &mut changes {
            if &change.key == key {
                change.before = before;
                change.after = after;
                replaced = true;
            }
        }
        if !replaced {
            changes.push(ChangeLengths {
                key: key.clone(),
                content_type: content_type.to_owned(),
                relationship_type: relationship_type.to_owned(),
                before,
                after,
            });
        }
        let dependents = self
            .dependents
            .iter()
            .map(dependent_lengths)
            .collect::<Result<Vec<_>>>()?;
        encoded_len_from_lengths(
            &self.before.order,
            &self.after.order,
            state_lengths(&self.before),
            state_lengths(&self.after),
            &changes,
            &dependents,
        )
    }

    fn version(&self) -> Version {
        if self.dependents.is_empty() {
            Version::Three
        } else {
            Version::Five
        }
    }
}

fn state_lengths(state: &State) -> StateLengths {
    StateLengths {
        root: state.root.len(),
        presentation_relationships: state.presentation_relationships.len(),
        content_types: state.content_types.len(),
    }
}

fn change_lengths(change: &Change) -> Result<ChangeLengths> {
    Ok(ChangeLengths {
        key: change.key.clone(),
        content_type: change.content_type.clone(),
        relationship_type: change.relationship_type.clone(),
        before: change.before.as_ref().map(image_lengths),
        after: change.after.as_ref().map(image_lengths),
    })
}

fn dependent_lengths(dependent: &Dependent) -> Result<DependentLengths> {
    Ok(DependentLengths {
        part_uri: dependent.part_uri.clone(),
        content_type: dependent.content_type.clone(),
        before: dependent.before.as_ref().map(image_lengths),
        after: dependent.after.as_ref().map(image_lengths),
    })
}

pub(super) fn encoded_len_from_lengths(
    before_order: &[StableSlideKey],
    after_order: &[StableSlideKey],
    before: StateLengths,
    after: StateLengths,
    changes: &[ChangeLengths],
    dependents: &[DependentLengths],
) -> Result<usize> {
    let mut total = HEADER_BYTES
        .checked_add(CHECKSUM_BYTES)
        .ok_or_else(|| invalid("wire length overflows"))?;
    for order_key in before_order.iter().chain(after_order.iter()) {
        add_len(&mut total, key_len(order_key)?)?;
    }
    let before_payload = before
        .root
        .checked_add(before.presentation_relationships)
        .and_then(|length| length.checked_add(before.content_types))
        .ok_or_else(|| invalid("state payload length overflows"))?;
    let after_payload = after
        .root
        .checked_add(after.presentation_relationships)
        .and_then(|length| length.checked_add(after.content_types))
        .ok_or_else(|| invalid("state payload length overflows"))?;
    add_len(&mut total, before_payload)?;
    add_len(&mut total, after_payload)?;
    for change in changes {
        add_len(
            &mut total,
            change_len_from_lengths(
                &change.key,
                &change.content_type,
                &change.relationship_type,
                change.before,
                change.after,
            )?,
        )?;
    }
    if !dependents.is_empty() {
        add_len(&mut total, 4)?;
        for dependent in dependents {
            add_len(&mut total, dependent_len_from_lengths(dependent)?)?;
        }
    }
    if total > MAX_WIRE_BYTES {
        return Err(wire_limit(total));
    }
    Ok(total)
}

fn dependent_len_from_lengths(dependent: &DependentLengths) -> Result<usize> {
    if dependent.part_uri.as_str().len() > u32::MAX as usize
        || dependent.content_type.len() > u32::MAX as usize
    {
        return Err(invalid("dependent Part exceeds wire width"));
    }
    let before_len = image_lengths_len(dependent.before)?;
    let after_len = image_lengths_len(dependent.after)?;
    8usize
        .checked_add(dependent.part_uri.as_str().len())
        .and_then(|length| length.checked_add(dependent.content_type.len()))
        .and_then(|length| length.checked_add(before_len))
        .and_then(|length| length.checked_add(after_len))
        .ok_or_else(|| invalid("dependent length overflows"))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Version {
    Three,
    Five,
}

#[derive(Clone, Copy)]
struct Header {
    before_count: usize,
    after_count: usize,
    change_count: usize,
    before: StateHeader,
    after: StateHeader,
}

#[derive(Clone, Copy)]
struct StateHeader {
    root_len: usize,
    relationships_len: usize,
    content_types_len: usize,
    topology_digest: [u8; 32],
}

fn write_header(output: &mut Vec<u8>, data: &Data) -> Result<()> {
    match data.version() {
        Version::Three => {
            output.extend_from_slice(MAGIC);
            output.extend_from_slice(&VERSION.to_le_bytes());
        },
        Version::Five => {
            output.extend_from_slice(DEPENDENT_MAGIC);
            output.extend_from_slice(&DEPENDENT_VERSION.to_le_bytes());
        },
    }
    put_u32(output, data.before.order.len())?;
    put_u32(output, data.after.order.len())?;
    put_u32(output, data.changes.len())?;
    write_state_header(output, &data.before)?;
    write_state_header(output, &data.after)?;
    Ok(())
}

fn write_state_header(output: &mut Vec<u8>, state: &State) -> Result<()> {
    put_u64(output, state.root.len())?;
    put_u64(output, state.presentation_relationships.len())?;
    put_u64(output, state.content_types.len())?;
    output.extend_from_slice(&state.topology_digest);
    Ok(())
}

fn write_order(output: &mut Vec<u8>, order: &[StableSlideKey]) -> Result<()> {
    for key in order {
        write_key(output, key)?;
    }
    Ok(())
}

fn write_state_payload(output: &mut Vec<u8>, state: &State) {
    output.extend_from_slice(state.root.as_slice());
    output.extend_from_slice(state.presentation_relationships.as_slice());
    output.extend_from_slice(state.content_types.as_slice());
}

fn write_key(output: &mut Vec<u8>, key: &StableSlideKey) -> Result<()> {
    put_u32(output, key.part_uri.as_str().len())?;
    output.extend_from_slice(key.part_uri.as_str().as_bytes());
    output.extend_from_slice(&key.slide_id.to_le_bytes());
    put_u32(output, key.relationship_id.len())?;
    output.extend_from_slice(key.relationship_id.as_bytes());
    Ok(())
}

fn write_string(output: &mut Vec<u8>, value: &str) -> Result<()> {
    put_u32(output, value.len())?;
    output.extend_from_slice(value.as_bytes());
    Ok(())
}

fn write_image(output: &mut Vec<u8>, image: Option<&PartImage>) -> Result<()> {
    match image {
        None => output.push(0),
        Some(image) => {
            output.push(1);
            put_u64(output, image.xml.len())?;
            output.extend_from_slice(image.xml.as_slice());
            match &image.relationships {
                None => output.push(0),
                Some(relationships) => {
                    output.push(1);
                    put_u64(output, relationships.len())?;
                    output.extend_from_slice(relationships.as_slice());
                },
            }
        },
    }
    Ok(())
}

fn read_header(input: &mut Input<'_>, limits: SourceLimits, version: Version) -> Result<Header> {
    let (magic, number) = match version {
        Version::Three => (MAGIC, VERSION),
        Version::Five => (DEPENDENT_MAGIC, DEPENDENT_VERSION),
    };
    if input.take(magic.len())? != magic {
        return Err(invalid("unsupported wire version"));
    }
    if input.u32()? != number {
        return Err(invalid("unsupported wire version"));
    }
    let before_count = input.count()?;
    let after_count = input.count()?;
    let change_count = input.count()?;
    if before_count > limits.max_parts() || after_count > limits.max_parts() {
        return Err(source_limit(
            SourceLimitKind::Parts,
            before_count.max(after_count) as u64,
            limits.max_parts() as u64,
        ));
    }
    if change_count > MAX_SOURCE_BACKED_SLIDE_BATCH {
        return Err(limit(
            crate::error::CodecLimitKind::Slides,
            change_count,
            MAX_SOURCE_BACKED_SLIDE_BATCH,
        ));
    }
    let before = read_state_header(input, limits)?;
    let after = read_state_header(input, limits)?;
    let minimum = minimum_body_len(before_count, after_count, change_count, &before, &after)?;
    if minimum > input.bytes.len().saturating_sub(input.offset) {
        return Err(invalid("counts exceed remaining framing"));
    }
    Ok(Header {
        before_count,
        after_count,
        change_count,
        before,
        after,
    })
}

fn read_state_header(input: &mut Input<'_>, limits: SourceLimits) -> Result<StateHeader> {
    let root_len = input.length()?;
    check_payload_len(
        root_len,
        limits.max_part_bytes(),
        SourceLimitKind::PartBytes,
        "presentation root bytes",
    )?;
    let relationships_len = input.length()?;
    check_payload_len(
        relationships_len,
        limits.max_relationship_xml_bytes() as u64,
        SourceLimitKind::RelationshipXmlBytes,
        "presentation relationship bytes",
    )?;
    let content_types_len = input.length()?;
    check_payload_len(
        content_types_len,
        limits.max_content_types_bytes() as u64,
        SourceLimitKind::ContentTypesBytes,
        "content-types bytes",
    )?;
    let topology_digest = input.digest()?;
    Ok(StateHeader {
        root_len,
        relationships_len,
        content_types_len,
        topology_digest,
    })
}

fn preflight_body(
    input: &mut Input<'_>,
    header: &Header,
    limits: SourceLimits,
    version: Version,
) -> Result<()> {
    for _ in 0..header.before_count {
        scan_key(input, limits)?;
    }
    for _ in 0..header.after_count {
        scan_key(input, limits)?;
    }

    let mut before_parts = u64::try_from(header.before.root_len)
        .map_err(|_| invalid("before part aggregate overflows"))?;
    let mut after_parts = u64::try_from(header.after.root_len)
        .map_err(|_| invalid("after part aggregate overflows"))?;
    let mut before_relationships = u64::try_from(header.before.relationships_len)
        .map_err(|_| invalid("before relationship aggregate overflows"))?;
    let mut after_relationships = u64::try_from(header.after.relationships_len)
        .map_err(|_| invalid("after relationship aggregate overflows"))?;
    input.skip(header.before.root_len)?;
    input.skip(header.before.relationships_len)?;
    input.skip(header.before.content_types_len)?;
    input.skip(header.after.root_len)?;
    input.skip(header.after.relationships_len)?;
    input.skip(header.after.content_types_len)?;

    for _ in 0..header.change_count {
        scan_key(input, limits)?;
        scan_string(
            input,
            string_limit(limits.max_xml_attribute_bytes()),
            "content type",
        )?;
        scan_string(
            input,
            string_limit(
                limits
                    .max_xml_attribute_bytes()
                    .min(limits.max_relationship_target_bytes()),
            ),
            "relationship type",
        )?;
        scan_image(input, limits, &mut before_parts, &mut before_relationships)?;
        scan_image(input, limits, &mut after_parts, &mut after_relationships)?;
    }
    if version == Version::Five {
        let count = input.count()?;
        if count > MAX_DEPENDENTS {
            return Err(limit(
                crate::error::CodecLimitKind::Parts,
                count,
                MAX_DEPENDENTS,
            ));
        }
        for _ in 0..count {
            scan_string(
                input,
                string_limit_u64(limits.max_archive_member_name_bytes()),
                "dependent URI",
            )?;
            scan_string(
                input,
                string_limit(limits.max_xml_attribute_bytes()),
                "content type",
            )?;
            scan_image(input, limits, &mut before_parts, &mut before_relationships)?;
            scan_image(input, limits, &mut after_parts, &mut after_relationships)?;
        }
    }
    if before_parts > limits.max_total_part_bytes() {
        return Err(source_limit(
            SourceLimitKind::TotalPartBytes,
            before_parts,
            limits.max_total_part_bytes(),
        ));
    }
    if after_parts > limits.max_total_part_bytes() {
        return Err(source_limit(
            SourceLimitKind::TotalPartBytes,
            after_parts,
            limits.max_total_part_bytes(),
        ));
    }
    if before_relationships > limits.max_total_relationship_xml_bytes() as u64 {
        return Err(source_limit(
            SourceLimitKind::TotalRelationshipXmlBytes,
            before_relationships,
            limits.max_total_relationship_xml_bytes() as u64,
        ));
    }
    if after_relationships > limits.max_total_relationship_xml_bytes() as u64 {
        return Err(source_limit(
            SourceLimitKind::TotalRelationshipXmlBytes,
            after_relationships,
            limits.max_total_relationship_xml_bytes() as u64,
        ));
    }
    if input.offset != input.bytes.len() {
        return Err(invalid("trailing bytes"));
    }
    Ok(())
}

fn scan_key(input: &mut Input<'_>, limits: SourceLimits) -> Result<()> {
    let uri = scan_string(
        input,
        string_limit_u64(limits.max_archive_member_name_bytes()),
        "stable key URI",
    )?;
    let _slide_id = input.u32()?;
    let relationship = scan_string(
        input,
        string_limit(
            limits
                .max_xml_attribute_bytes()
                .min(limits.max_relationship_target_bytes()),
        ),
        "stable key relationship ID",
    )?;
    if uri.is_empty() || relationship.is_empty() {
        return Err(invalid("stable key strings must not be empty"));
    }
    Ok(())
}

fn scan_string<'a>(input: &mut Input<'a>, maximum: u64, name: &'static str) -> Result<&'a str> {
    let length = input.count()?;
    if u64::try_from(length).map_err(|_| invalid("string length overflows"))? > maximum {
        return Err(source_limit(
            SourceLimitKind::XmlAttributeBytes,
            length as u64,
            maximum,
        ));
    }
    let bytes = input.take(length)?;
    std::str::from_utf8(bytes).map_err(|_| {
        invalid(match name {
            "stable key URI" => "stable key URI is not UTF-8",
            "stable key relationship ID" => "stable key relationship ID is not UTF-8",
            "content type" => "content type is not UTF-8",
            "relationship type" => "relationship type is not UTF-8",
            _ => "wire string is not UTF-8",
        })
    })
}

fn scan_image(
    input: &mut Input<'_>,
    limits: SourceLimits,
    parts: &mut u64,
    relationships: &mut u64,
) -> Result<()> {
    match input.byte()? {
        0 => Ok(()),
        1 => {
            let xml_len = input.length()?;
            check_payload_len(
                xml_len,
                limits.max_part_bytes(),
                SourceLimitKind::PartBytes,
                "slide part bytes",
            )?;
            *parts = parts
                .checked_add(xml_len as u64)
                .ok_or_else(|| invalid("part aggregate overflows"))?;
            input.skip(xml_len)?;
            match input.byte()? {
                0 => Ok(()),
                1 => {
                    let relationship_len = input.length()?;
                    check_payload_len(
                        relationship_len,
                        limits.max_relationship_xml_bytes() as u64,
                        SourceLimitKind::RelationshipXmlBytes,
                        "slide relationship bytes",
                    )?;
                    *relationships = relationships
                        .checked_add(relationship_len as u64)
                        .ok_or_else(|| invalid("relationship aggregate overflows"))?;
                    input.skip(relationship_len)?;
                    Ok(())
                },
                _ => Err(invalid("invalid relationship image marker")),
            }
        },
        _ => Err(invalid("invalid part image marker")),
    }
}

fn read_order(
    input: &mut Input<'_>,
    count: usize,
    limits: SourceLimits,
) -> Result<Vec<StableSlideKey>> {
    let mut order = Vec::new();
    order.try_reserve_exact(count).map_err(allocation)?;
    let mut uris = HashSet::new();
    let mut slide_ids = HashSet::new();
    let mut relationship_ids = HashSet::new();
    uris.try_reserve(count).map_err(allocation)?;
    slide_ids.try_reserve(count).map_err(allocation)?;
    relationship_ids.try_reserve(count).map_err(allocation)?;
    for _ in 0..count {
        let key = read_key(input, limits)?;
        insert_identity(&key, &mut uris, &mut slide_ids, &mut relationship_ids)?;
        order.push(key);
    }
    Ok(order)
}

fn read_key(input: &mut Input<'_>, limits: SourceLimits) -> Result<StableSlideKey> {
    let uri_len = input.count()?;
    check_string_length(
        uri_len,
        string_limit_u64(limits.max_archive_member_name_bytes()),
        "stable key URI",
    )?;
    let uri_bytes = input.take(uri_len)?;
    let uri_text =
        std::str::from_utf8(uri_bytes).map_err(|_| invalid("stable key URI is not UTF-8"))?;
    if uri_text.is_empty() {
        return Err(invalid("stable key URI is empty"));
    }
    let uri = PackURI::new(uri_text.to_owned()).map_err(|_| invalid("invalid stable key URI"))?;
    let slide_id = input.u32()?;
    let relationship_len = input.count()?;
    check_string_length(
        relationship_len,
        string_limit(
            limits
                .max_xml_attribute_bytes()
                .min(limits.max_relationship_target_bytes()),
        ),
        "stable key relationship ID",
    )?;
    let relationship_bytes = input.take(relationship_len)?;
    let relationship_id = std::str::from_utf8(relationship_bytes)
        .map_err(|_| invalid("stable key relationship ID is not UTF-8"))?;
    if relationship_id.is_empty() {
        return Err(invalid("stable key relationship ID is empty"));
    }
    Ok(stable_slide_key(uri, slide_id, relationship_id.to_owned()))
}

fn read_state(
    input: &mut Input<'_>,
    header: StateHeader,
    order: Vec<StableSlideKey>,
    limits: SourceLimits,
) -> Result<State> {
    let root = copy(input.take(header.root_len)?)?;
    let presentation_relationships = copy(input.take(header.relationships_len)?)?;
    let content_types = copy(input.take(header.content_types_len)?)?;
    let state = State {
        root,
        order: order.into(),
        topology_digest: header.topology_digest,
        presentation_relationships,
        content_types,
    };
    validate_state(&state, limits)?;
    Ok(state)
}

fn read_changes(input: &mut Input<'_>, count: usize, limits: SourceLimits) -> Result<Vec<Change>> {
    let mut changes = Vec::new();
    changes.try_reserve_exact(count).map_err(allocation)?;
    let mut uris = HashSet::new();
    let mut slide_ids = HashSet::new();
    let mut relationship_ids = HashSet::new();
    uris.try_reserve(count).map_err(allocation)?;
    slide_ids.try_reserve(count).map_err(allocation)?;
    relationship_ids.try_reserve(count).map_err(allocation)?;
    for _ in 0..count {
        let key = read_key(input, limits)?;
        insert_identity(&key, &mut uris, &mut slide_ids, &mut relationship_ids)?;
        let content_type = read_string(
            input,
            string_limit(limits.max_xml_attribute_bytes()),
            "content type",
        )?;
        let relationship_type = read_string(
            input,
            string_limit(
                limits
                    .max_xml_attribute_bytes()
                    .min(limits.max_relationship_target_bytes()),
            ),
            "relationship type",
        )?;
        if content_type.is_empty() || relationship_type.is_empty() {
            return Err(invalid("resource change strings must not be empty"));
        }
        let before = read_image(input, limits)?;
        let after = read_image(input, limits)?;
        changes.push(Change {
            key,
            content_type,
            relationship_type,
            before,
            after,
        });
    }
    Ok(changes)
}

fn read_dependents(input: &mut Input<'_>, limits: SourceLimits) -> Result<Vec<Dependent>> {
    let count = input.count()?;
    if count > MAX_DEPENDENTS {
        return Err(limit(
            crate::error::CodecLimitKind::Parts,
            count,
            MAX_DEPENDENTS,
        ));
    }
    let mut dependents = Vec::new();
    dependents.try_reserve_exact(count).map_err(allocation)?;
    for _ in 0..count {
        let uri = read_string(
            input,
            string_limit_u64(limits.max_archive_member_name_bytes()),
            "dependent URI",
        )?;
        let part_uri = PackURI::new(uri).map_err(|_| invalid("invalid dependent Part URI"))?;
        let content_type = read_string(
            input,
            string_limit(limits.max_xml_attribute_bytes()),
            "content type",
        )?;
        let before = read_image(input, limits)?;
        let after = read_image(input, limits)?;
        dependents.push(Dependent {
            part_uri,
            content_type,
            before,
            after,
        });
    }
    Ok(dependents)
}

fn read_string(input: &mut Input<'_>, maximum: u64, name: &'static str) -> Result<String> {
    let length = input.count()?;
    check_string_length(length, maximum, name)?;
    let bytes = input.take(length)?;
    let text = std::str::from_utf8(bytes).map_err(|_| {
        invalid(match name {
            "content type" => "content type is not UTF-8",
            "relationship type" => "relationship type is not UTF-8",
            _ => "wire string is not UTF-8",
        })
    })?;
    let mut value = String::new();
    value.try_reserve_exact(length).map_err(allocation)?;
    value.push_str(text);
    Ok(value)
}

fn read_image(input: &mut Input<'_>, limits: SourceLimits) -> Result<Option<PartImage>> {
    match input.byte()? {
        0 => Ok(None),
        1 => {
            let xml_len = input.length()?;
            check_payload_len(
                xml_len,
                limits.max_part_bytes(),
                SourceLimitKind::PartBytes,
                "slide part bytes",
            )?;
            let xml = copy(input.take(xml_len)?)?;
            let relationships = match input.byte()? {
                0 => None,
                1 => {
                    let length = input.length()?;
                    check_payload_len(
                        length,
                        limits.max_relationship_xml_bytes() as u64,
                        SourceLimitKind::RelationshipXmlBytes,
                        "slide relationship bytes",
                    )?;
                    Some(copy(input.take(length)?)?)
                },
                _ => return Err(invalid("invalid relationship image marker")),
            };
            Ok(Some(PartImage { xml, relationships }))
        },
        _ => Err(invalid("invalid part image marker")),
    }
}

fn validate_data(data: &Data, limits: SourceLimits) -> Result<()> {
    if data.changes.len() > MAX_SOURCE_BACKED_SLIDE_BATCH {
        return Err(limit(
            crate::error::CodecLimitKind::Slides,
            data.changes.len(),
            MAX_SOURCE_BACKED_SLIDE_BATCH,
        ));
    }
    validate_state(&data.before, limits)?;
    validate_state(&data.after, limits)?;
    let before_set = key_set(&data.before.order)?;
    let after_set = key_set(&data.after.order)?;
    let mut change_uris = HashSet::new();
    let mut change_slide_ids = HashSet::new();
    let mut change_relationship_ids = HashSet::new();
    change_uris
        .try_reserve(data.changes.len())
        .map_err(allocation)?;
    change_slide_ids
        .try_reserve(data.changes.len())
        .map_err(allocation)?;
    change_relationship_ids
        .try_reserve(data.changes.len())
        .map_err(allocation)?;
    for change in data.changes.iter() {
        validate_key(&change.key, limits)?;
        insert_identity(
            &change.key,
            &mut change_uris,
            &mut change_slide_ids,
            &mut change_relationship_ids,
        )?;
        validate_string(
            &change.content_type,
            string_limit(limits.max_xml_attribute_bytes()),
            "content type",
        )?;
        validate_string(
            &change.relationship_type,
            string_limit(
                limits
                    .max_xml_attribute_bytes()
                    .min(limits.max_relationship_target_bytes()),
            ),
            "relationship type",
        )?;
        if change.content_type.is_empty() || change.relationship_type.is_empty() {
            return Err(invalid("resource change strings must not be empty"));
        }
        let before_member = contains_key(&before_set, &change.key);
        let after_member = contains_key(&after_set, &change.key);
        if !before_member && !after_member {
            return Err(invalid("resource change references an unknown stable key"));
        }
        if before_member != change.before.is_some() || after_member != change.after.is_some() {
            return Err(invalid(
                "resource image presence disagrees with order membership",
            ));
        }
        if change.before == change.after {
            return Err(invalid("resource change is an identical no-op"));
        }
        if let Some(image) = change.before.as_ref() {
            validate_image(image, limits)?;
        }
        if let Some(image) = change.after.as_ref() {
            validate_image(image, limits)?;
        }
    }

    for key in data.before.order.iter() {
        if !contains_key(&after_set, key) {
            require_change_presence(&data.changes, key, true, false)?;
        }
    }
    for key in data.after.order.iter() {
        if !contains_key(&before_set, key) {
            require_change_presence(&data.changes, key, false, true)?;
        }
    }
    validate_dependents(data, &change_uris, limits)?;
    validate_change_aggregates(data, limits)?;
    Ok(())
}

/// Dependent Parts are unique, sorted by folded name, distinct from every
/// slide the patch changes, and each one changes presence: the wire carries
/// only Parts that a slide removal deletes or its inverse restores.
fn validate_dependents(
    data: &Data,
    change_uris: &HashSet<String>,
    limits: SourceLimits,
) -> Result<()> {
    if data.dependents.len() > MAX_DEPENDENTS {
        return Err(limit(
            crate::error::CodecLimitKind::Parts,
            data.dependents.len(),
            MAX_DEPENDENTS,
        ));
    }
    let mut previous: Option<String> = None;
    for dependent in data.dependents.iter() {
        let uri = dependent.part_uri.as_str();
        if uri.len() > u32::MAX as usize {
            return Err(invalid("dependent URI exceeds wire width"));
        }
        check_string_length(
            uri.len(),
            string_limit_u64(limits.max_archive_member_name_bytes()),
            "dependent URI",
        )?;
        validate_string(
            &dependent.content_type,
            string_limit(limits.max_xml_attribute_bytes()),
            "content type",
        )?;
        if dependent.content_type.is_empty() {
            return Err(invalid("dependent content type must not be empty"));
        }
        let folded = folded_ascii(uri)?;
        if change_uris.contains(&folded) {
            return Err(invalid("a dependent Part is also a changed slide"));
        }
        if previous
            .as_ref()
            .is_some_and(|previous| previous >= &folded)
        {
            return Err(invalid("dependent Parts are not sorted and unique"));
        }
        previous = Some(folded);
        if dependent.before.is_some() == dependent.after.is_some() {
            return Err(invalid("a dependent Part must be removed or restored"));
        }
        for image in [dependent.before.as_ref(), dependent.after.as_ref()]
            .into_iter()
            .flatten()
        {
            validate_image(image, limits)?;
        }
    }
    Ok(())
}

fn validate_state(state: &State, limits: SourceLimits) -> Result<()> {
    if state.root.is_empty() {
        return Err(invalid("source topology roots must not be empty"));
    }
    if state.order.len() > limits.max_parts() {
        return Err(source_limit(
            SourceLimitKind::Parts,
            state.order.len() as u64,
            limits.max_parts() as u64,
        ));
    }
    check_payload_len(
        state.root.len(),
        limits.max_part_bytes(),
        SourceLimitKind::PartBytes,
        "presentation root bytes",
    )?;
    check_payload_len(
        state.presentation_relationships.len(),
        limits.max_relationship_xml_bytes() as u64,
        SourceLimitKind::RelationshipXmlBytes,
        "presentation relationship bytes",
    )?;
    check_payload_len(
        state.content_types.len(),
        limits.max_content_types_bytes() as u64,
        SourceLimitKind::ContentTypesBytes,
        "content-types bytes",
    )?;
    let mut uris = HashSet::new();
    let mut slide_ids = HashSet::new();
    let mut relationship_ids = HashSet::new();
    uris.try_reserve(state.order.len()).map_err(allocation)?;
    slide_ids
        .try_reserve(state.order.len())
        .map_err(allocation)?;
    relationship_ids
        .try_reserve(state.order.len())
        .map_err(allocation)?;
    for key in state.order.iter() {
        validate_key(key, limits)?;
        insert_identity(key, &mut uris, &mut slide_ids, &mut relationship_ids)?;
    }
    Ok(())
}

fn validate_key(key: &StableSlideKey, limits: SourceLimits) -> Result<()> {
    let uri = key.part_uri.as_str();
    if uri.is_empty() {
        return Err(invalid("stable key URI is empty"));
    }
    if uri.len() > u32::MAX as usize {
        return Err(invalid("stable key URI exceeds wire width"));
    }
    check_string_length(
        uri.len(),
        string_limit_u64(limits.max_archive_member_name_bytes()),
        "stable key URI",
    )?;
    if key.relationship_id.is_empty() {
        return Err(invalid("stable key relationship ID is empty"));
    }
    if key.relationship_id.len() > u32::MAX as usize {
        return Err(invalid("stable key relationship ID exceeds wire width"));
    }
    check_string_length(
        key.relationship_id.len(),
        string_limit(
            limits
                .max_xml_attribute_bytes()
                .min(limits.max_relationship_target_bytes()),
        ),
        "stable key relationship ID",
    )?;
    Ok(())
}

fn validate_string(value: &str, maximum: u64, name: &'static str) -> Result<()> {
    if value.len() > u32::MAX as usize {
        return Err(invalid(match name {
            "content type" => "content type exceeds wire width",
            "relationship type" => "relationship type exceeds wire width",
            _ => "wire string exceeds wire width",
        }));
    }
    check_string_length(value.len(), maximum, name)
}

fn validate_image(image: &PartImage, limits: SourceLimits) -> Result<()> {
    check_payload_len(
        image.xml.len(),
        limits.max_part_bytes(),
        SourceLimitKind::PartBytes,
        "slide part bytes",
    )?;
    if let Some(relationships) = image.relationships.as_ref() {
        check_payload_len(
            relationships.len(),
            limits.max_relationship_xml_bytes() as u64,
            SourceLimitKind::RelationshipXmlBytes,
            "slide relationship bytes",
        )?;
    }
    Ok(())
}

fn validate_change_aggregates(data: &Data, limits: SourceLimits) -> Result<()> {
    for (state, before) in [(&data.before, true), (&data.after, false)] {
        let mut parts =
            u64::try_from(state.root.len()).map_err(|_| invalid("part aggregate overflows"))?;
        let mut relationships = u64::try_from(state.presentation_relationships.len())
            .map_err(|_| invalid("relationship aggregate overflows"))?;
        let images = data
            .changes
            .iter()
            .map(|change| (&change.before, &change.after))
            .chain(
                data.dependents
                    .iter()
                    .map(|dependent| (&dependent.before, &dependent.after)),
            );
        for (before_image, after_image) in images {
            let image = if before {
                before_image.as_ref()
            } else {
                after_image.as_ref()
            };
            if let Some(image) = image {
                parts = parts
                    .checked_add(image.xml.len() as u64)
                    .ok_or_else(|| invalid("part aggregate overflows"))?;
                if let Some(relationships_image) = image.relationships.as_ref() {
                    relationships = relationships
                        .checked_add(relationships_image.len() as u64)
                        .ok_or_else(|| invalid("relationship aggregate overflows"))?;
                }
            }
        }
        if parts > limits.max_total_part_bytes() {
            return Err(source_limit(
                SourceLimitKind::TotalPartBytes,
                parts,
                limits.max_total_part_bytes(),
            ));
        }
        if relationships > limits.max_total_relationship_xml_bytes() as u64 {
            return Err(source_limit(
                SourceLimitKind::TotalRelationshipXmlBytes,
                relationships,
                limits.max_total_relationship_xml_bytes() as u64,
            ));
        }
    }
    Ok(())
}

fn require_change_presence(
    changes: &[Change],
    key: &StableSlideKey,
    before: bool,
    after: bool,
) -> Result<()> {
    let Some(change) = changes.iter().find(|change| &change.key == key) else {
        return Err(invalid("added or removed key lacks a resource change"));
    };
    if change.before.is_some() != before || change.after.is_some() != after {
        return Err(invalid(
            "added or removed key has incorrect resource change",
        ));
    }
    Ok(())
}

fn key_set(order: &[StableSlideKey]) -> Result<HashSet<&StableSlideKey>> {
    let mut set = HashSet::new();
    set.try_reserve(order.len()).map_err(allocation)?;
    set.extend(order.iter());
    Ok(set)
}

fn contains_key(set: &HashSet<&StableSlideKey>, key: &StableSlideKey) -> bool {
    set.contains(key)
}

fn insert_identity(
    key: &StableSlideKey,
    uris: &mut HashSet<String>,
    slide_ids: &mut HashSet<u32>,
    relationship_ids: &mut HashSet<String>,
) -> Result<()> {
    let uri = folded_ascii(key.part_uri.as_str())?;
    if !uris.insert(uri) {
        return Err(invalid("stable order repeats a case-equivalent part URI"));
    }
    if !slide_ids.insert(key.slide_id) {
        return Err(invalid("stable order repeats a slide ID"));
    }
    if !relationship_ids.insert(key.relationship_id.clone()) {
        return Err(invalid("stable order repeats a relationship ID"));
    }
    Ok(())
}

fn folded_ascii(value: &str) -> Result<String> {
    let mut folded = String::new();
    folded.try_reserve_exact(value.len()).map_err(allocation)?;
    for byte in value.bytes() {
        folded.push(byte.to_ascii_lowercase() as char);
    }
    Ok(folded)
}

fn change_len_from_lengths(
    key: &StableSlideKey,
    content_type: &str,
    relationship_type: &str,
    before: Option<ImageLengths>,
    after: Option<ImageLengths>,
) -> Result<usize> {
    let mut total = key_len(key)?;
    total = total
        .checked_add(4)
        .and_then(|length| length.checked_add(content_type.len()))
        .and_then(|length| length.checked_add(4))
        .and_then(|length| length.checked_add(relationship_type.len()))
        .ok_or_else(|| invalid("change length overflows"))?;
    let before_len = image_lengths_len(before)?;
    let after_len = image_lengths_len(after)?;
    total = total
        .checked_add(before_len)
        .and_then(|length| length.checked_add(after_len))
        .ok_or_else(|| invalid("change length overflows"))?;
    Ok(total)
}

fn image_lengths(image: &PartImage) -> ImageLengths {
    ImageLengths {
        xml: image.xml.len(),
        relationships: image.relationships.as_ref().map(|bytes| bytes.len()),
    }
}

fn image_lengths_len(image: Option<ImageLengths>) -> Result<usize> {
    let mut total = 1usize;
    if let Some(image) = image {
        total = total
            .checked_add(8)
            .and_then(|length| length.checked_add(image.xml))
            .and_then(|length| length.checked_add(1))
            .ok_or_else(|| invalid("image length overflows"))?;
        if let Some(relationships) = image.relationships {
            total = total
                .checked_add(8)
                .and_then(|length| length.checked_add(relationships))
                .ok_or_else(|| invalid("image length overflows"))?;
        }
    }
    Ok(total)
}

fn key_len(key: &StableSlideKey) -> Result<usize> {
    if key.part_uri.as_str().len() > u32::MAX as usize
        || key.relationship_id.len() > u32::MAX as usize
    {
        return Err(invalid("stable key exceeds wire width"));
    }
    KEY_FIXED_BYTES
        .checked_add(key.part_uri.as_str().len())
        .and_then(|length| length.checked_add(key.relationship_id.len()))
        .ok_or_else(|| invalid("stable key length overflows"))
}

fn minimum_body_len(
    before_count: usize,
    after_count: usize,
    change_count: usize,
    before: &StateHeader,
    after: &StateHeader,
) -> Result<usize> {
    let keys = before_count
        .checked_add(after_count)
        .and_then(|count| count.checked_mul(KEY_FIXED_BYTES))
        .ok_or_else(|| invalid("catalog framing overflows"))?;
    let changes = change_count
        .checked_mul(CHANGE_FIXED_BYTES)
        .ok_or_else(|| invalid("change framing overflows"))?;
    HEADER_BYTES
        .checked_add(keys)
        .and_then(|length| length.checked_add(changes))
        .and_then(|length| length.checked_add(before.root_len))
        .and_then(|length| length.checked_add(before.relationships_len))
        .and_then(|length| length.checked_add(before.content_types_len))
        .and_then(|length| length.checked_add(after.root_len))
        .and_then(|length| length.checked_add(after.relationships_len))
        .and_then(|length| length.checked_add(after.content_types_len))
        .ok_or_else(|| invalid("wire framing overflows"))
        .and_then(|length| {
            // `length` includes the header already consumed by `Input`; the
            // comparison in `read_header` is against the remaining body.
            length
                .checked_sub(HEADER_BYTES)
                .ok_or_else(|| invalid("wire framing underflows"))
        })
}

fn add_len(total: &mut usize, additional: usize) -> Result<()> {
    *total = total
        .checked_add(additional)
        .ok_or_else(|| invalid("wire length overflows"))?;
    Ok(())
}

fn put_u32(output: &mut Vec<u8>, value: usize) -> Result<()> {
    output.extend_from_slice(
        &u32::try_from(value)
            .map_err(|_| invalid("integer exceeds wire width"))?
            .to_le_bytes(),
    );
    Ok(())
}

fn put_u64(output: &mut Vec<u8>, value: usize) -> Result<()> {
    output.extend_from_slice(
        &u64::try_from(value)
            .map_err(|_| invalid("length exceeds wire width"))?
            .to_le_bytes(),
    );
    Ok(())
}

fn checksum(bytes: &[u8], version: Version) -> [u8; 32] {
    let mut hash = Sha256::new();
    hash.update(match version {
        Version::Three => CHECKSUM_DOMAIN,
        Version::Five => DEPENDENT_CHECKSUM_DOMAIN,
    });
    hash.update(bytes);
    hash.finalize().into()
}

fn copy(bytes: &[u8]) -> Result<Arc<Vec<u8>>> {
    let mut output = Vec::new();
    output.try_reserve_exact(bytes.len()).map_err(allocation)?;
    output.extend_from_slice(bytes);
    Ok(Arc::new(output))
}

fn check_payload_len(
    length: usize,
    maximum: u64,
    kind: SourceLimitKind,
    resource: &'static str,
) -> Result<()> {
    if u64::try_from(length).map_err(|_| invalid("payload length overflows"))?
        > MAX_WIRE_BYTES as u64
    {
        return Err(wire_limit(length));
    }
    if length as u64 > maximum {
        return Err(source_limit(kind, length as u64, maximum));
    }
    let _ = resource;
    Ok(())
}

fn check_string_length(length: usize, maximum: u64, name: &'static str) -> Result<()> {
    if length > u32::MAX as usize {
        return Err(invalid(match name {
            "stable key URI" => "stable key URI exceeds wire width",
            "stable key relationship ID" => "stable key relationship ID exceeds wire width",
            "content type" => "content type exceeds wire width",
            "relationship type" => "relationship type exceeds wire width",
            _ => "wire string exceeds wire width",
        }));
    }
    if length as u64 > MAX_WIRE_BYTES as u64 {
        return Err(wire_limit(length));
    }
    if length as u64 > maximum {
        return Err(source_limit(
            SourceLimitKind::XmlAttributeBytes,
            length as u64,
            maximum,
        ));
    }
    Ok(())
}

fn string_limit(value: usize) -> u64 {
    string_limit_u64(value as u64)
}

fn string_limit_u64(value: u64) -> u64 {
    value.min(MAX_WIRE_BYTES as u64)
}

fn wire_limit(observed: usize) -> Error {
    limit(
        crate::error::CodecLimitKind::SourceBytes,
        observed,
        MAX_WIRE_BYTES,
    )
}

fn source_limit(kind: SourceLimitKind, observed: u64, maximum: u64) -> Error {
    Error::Opc(litchi_opc::OpcError::Limit {
        kind,
        observed,
        maximum,
    })
}

fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(format!("source topology wire: {message}"))
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    Error::Allocation {
        resource: "source-backed topology wire",
        source,
    }
}

struct Input<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Input<'a> {
    fn take(&mut self, length: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(length)
            .ok_or_else(|| invalid("offset overflows"))?;
        let bytes = self
            .bytes
            .get(self.offset..end)
            .ok_or_else(|| invalid("truncated encoding"))?;
        self.offset = end;
        Ok(bytes)
    }

    fn skip(&mut self, length: usize) -> Result<()> {
        self.take(length).map(|_| ())
    }

    fn byte(&mut self) -> Result<u8> {
        self.take(1)?
            .first()
            .copied()
            .ok_or_else(|| invalid("byte width"))
    }

    fn u32(&mut self) -> Result<u32> {
        self.take(4)?
            .try_into()
            .map(u32::from_le_bytes)
            .map_err(|_| invalid("integer width"))
    }

    fn count(&mut self) -> Result<usize> {
        usize::try_from(self.u32()?).map_err(|_| invalid("integer overflows"))
    }

    fn length(&mut self) -> Result<usize> {
        usize::try_from(self.u64()?).map_err(|_| invalid("length overflows"))
    }

    fn u64(&mut self) -> Result<u64> {
        self.take(8)?
            .try_into()
            .map(u64::from_le_bytes)
            .map_err(|_| invalid("length width"))
    }

    fn digest(&mut self) -> Result<[u8; 32]> {
        self.take(32)?
            .try_into()
            .map_err(|_| invalid("digest width"))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn key(id: u32) -> StableSlideKey {
        stable_slide_key(
            PackURI::new(format!("/ppt/slides/slide{id}.xml")).expect("test URI"),
            id,
            format!("rId{id}"),
        )
    }

    fn image(xml: &[u8], relationships: Option<&[u8]>) -> PartImage {
        PartImage {
            xml: Arc::new(xml.to_vec()),
            relationships: relationships.map(|bytes| Arc::new(bytes.to_vec())),
        }
    }

    fn state(root: &[u8], order: Vec<StableSlideKey>, digest: u8) -> State {
        State {
            root: Arc::new(root.to_vec()),
            order: order.into(),
            topology_digest: [digest; 32],
            presentation_relationships: Arc::new(b"<Relationships/>".to_vec()),
            content_types: Arc::new(b"<Types/>".to_vec()),
        }
    }

    fn sample() -> Data {
        let first = key(1);
        let second = key(2);
        let third = key(3);
        let added = key(4);
        Data {
            dependents: Arc::from([]),
            before: state(
                b"<presentation before='1'/>",
                vec![first.clone(), second.clone(), third.clone()],
                1,
            ),
            after: state(
                b"<presentation after='1'/>",
                vec![first.clone(), third, added.clone()],
                2,
            ),
            changes: vec![
                Change {
                    key: first,
                    content_type:
                        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
                            .into(),
                    relationship_type:
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
                            .into(),
                    before: Some(image(b"<sld><txBody>before</txBody></sld>", Some(b""))),
                    after: Some(image(
                        b"<sld><txBody>after</txBody></sld>",
                        Some(b"<Relationships/>"),
                    )),
                },
                Change {
                    key: second,
                    content_type:
                        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
                            .into(),
                    relationship_type:
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
                            .into(),
                    before: Some(image(b"<removed/>", None)),
                    after: None,
                },
                Change {
                    key: added,
                    content_type:
                        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
                            .into(),
                    relationship_type:
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
                            .into(),
                    before: None,
                    after: Some(image(b"<added/>", Some(b""))),
                },
            ]
            .into(),
        }
    }

    #[test]
    fn replacement_length_matches_encoded_wire() {
        let data = sample();
        let change = &data.changes[0];
        let before = change.before.as_ref().map(image_lengths);
        let after = change.after.as_ref().map(image_lengths);
        let measured = data
            .encoded_len_with_replacement(
                &change.key,
                before,
                after,
                &change.content_type,
                &change.relationship_type,
            )
            .expect("length-only replacement");
        assert_eq!(measured, data.encoded_len().expect("encoded length"));

        let after_image = after.expect("after image");
        let replacement = ImageLengths {
            xml: after_image.xml + 17,
            relationships: after_image.relationships.map(|length| length + 3),
        };
        let expected = data.encoded_len().expect("encoded length") + 17 + 3;
        assert_eq!(
            data.encoded_len_with_replacement(
                &change.key,
                before,
                Some(replacement),
                &change.content_type,
                &change.relationship_type,
            )
            .expect("replacement length"),
            expected
        );
    }

    #[test]
    fn length_only_plan_includes_dependent_framing_and_checksum() {
        let mut data = sample();
        data.dependents = vec![Dependent {
            part_uri: PackURI::new("/ppt/media/image1.png").expect("dependent URI"),
            content_type: "image/png".into(),
            before: Some(image(b"PNG-bytes", Some(b"<Relationships/>"))),
            after: None,
        }]
        .into();
        let changes = data
            .changes
            .iter()
            .map(change_lengths)
            .collect::<Result<Vec<_>>>()
            .expect("change lengths");
        let dependents = data
            .dependents
            .iter()
            .map(dependent_lengths)
            .collect::<Result<Vec<_>>>()
            .expect("dependent lengths");
        let measured = encoded_len_from_lengths(
            &data.before.order,
            &data.after.order,
            state_lengths(&data.before),
            state_lengths(&data.after),
            &changes,
            &dependents,
        )
        .expect("length-only topology plan");
        assert_eq!(measured, data.encoded_len().expect("encoded topology wire"));
    }

    #[test]
    fn metadata_length_delta_is_admitted_before_structural_capture() {
        let data = sample();
        let mut candidate = data.clone();
        let mut relationships = candidate.after.presentation_relationships.as_ref().clone();
        relationships.extend_from_slice(b"<Relationship Id=\"rId-extra\"/>");
        candidate.after.presentation_relationships = Arc::new(relationships);
        let delta = candidate.after.presentation_relationships.len()
            - data.after.presentation_relationships.len();
        assert_eq!(
            candidate.encoded_len().expect("candidate wire length"),
            data.encoded_len().expect("base wire length") + delta
        );
    }

    fn empty_transition() -> Data {
        let added = key(7);
        Data {
            dependents: Arc::from([]),
            before: state(
                b"<p:presentation><p:sldIdLst/></p:presentation>",
                Vec::new(),
                9,
            ),
            after: state(
                b"<p:presentation><p:sldIdLst><p:sldId id=\"7\" r:id=\"rId7\"/></p:sldIdLst></p:presentation>",
                vec![added.clone()],
                10,
            ),
            changes: vec![Change {
                key: added,
                content_type: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
                    .into(),
                relationship_type:
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
                        .into(),
                before: None,
                after: Some(image(b"<p:sld/>", Some(b"<Relationships/>"))),
            }]
            .into(),
        }
    }

    fn resign(bytes: &mut [u8]) {
        let body = bytes.len() - CHECKSUM_BYTES;
        let version = if bytes.starts_with(DEPENDENT_MAGIC) {
            Version::Five
        } else {
            Version::Three
        };
        let digest = checksum(&bytes[..body], version);
        bytes[body..].copy_from_slice(&digest);
    }

    fn with_dependents() -> Data {
        let mut data = sample();
        data.dependents = vec![
            Dependent {
                part_uri: PackURI::new("/ppt/media/image1.png").expect("test URI"),
                content_type: "image/png".into(),
                before: Some(image(b"\x89PNG binary", None)),
                after: None,
            },
            Dependent {
                part_uri: PackURI::new("/ppt/notesSlides/notesSlide2.xml").expect("test URI"),
                content_type:
                    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
                        .into(),
                before: Some(image(b"<p:notes/>", Some(b"<Relationships/>"))),
                after: None,
            },
        ]
        .into();
        data
    }

    #[test]
    fn dependents_roundtrip_as_version_five_and_invert() {
        let data = with_dependents();
        let bytes = data.to_bytes().unwrap();
        assert!(bytes.starts_with(DEPENDENT_MAGIC));
        assert_eq!(bytes.len(), data.encoded_len().unwrap());
        let decoded = Data::from_bytes(&bytes, SourceLimits::default()).unwrap();
        assert_eq!(decoded, data);
        let inverse = decoded.inverse();
        assert!(
            inverse
                .dependents
                .iter()
                .all(|dependent| dependent.before.is_none())
        );
        assert_eq!(inverse.inverse(), data);
        // Without dependents the wire stays version 3.
        assert!(sample().to_bytes().unwrap().starts_with(MAGIC));
    }

    #[test]
    fn dependents_must_be_sorted_unique_and_change_presence() {
        let mut unsorted = with_dependents();
        let mut parts = unsorted.dependents.to_vec();
        parts.reverse();
        unsorted.dependents = parts.into();
        assert!(unsorted.to_bytes().is_err());

        let mut unchanged = with_dependents();
        let mut parts = unchanged.dependents.to_vec();
        parts[0].after = parts[0].before.clone();
        unchanged.dependents = parts.into();
        assert!(unchanged.to_bytes().is_err());

        // A version-5 frame must carry a dependent.
        let mut bytes = sample().to_bytes().unwrap();
        bytes[..8].copy_from_slice(DEPENDENT_MAGIC);
        bytes[8..12].copy_from_slice(&DEPENDENT_VERSION.to_le_bytes());
        let body = bytes.len() - CHECKSUM_BYTES;
        bytes.truncate(body);
        bytes.extend_from_slice(&0u32.to_le_bytes());
        let digest = checksum(&bytes, Version::Five);
        bytes.extend_from_slice(&digest);
        assert!(Data::from_bytes(&bytes, SourceLimits::default()).is_err());
    }

    #[test]
    fn roundtrip_multi_text_removal_and_addition() {
        let data = sample();
        let bytes = data.to_bytes().expect("encode");
        let decoded = Data::from_bytes(&bytes, SourceLimits::default()).expect("decode");
        assert_eq!(decoded, data);
        assert_eq!(decoded.to_bytes().expect("re-encode"), bytes);
    }

    #[test]
    fn inversion_is_exact_and_reversible() {
        let data = sample();
        let inverse = data.inverse();
        assert_eq!(inverse.before, data.after);
        assert_eq!(inverse.after, data.before);
        assert_eq!(inverse.inverse(), data);
    }

    #[test]
    fn empty_catalog_transition_roundtrips_and_inverts() {
        let data = empty_transition();
        let bytes = data.to_bytes().expect("encode empty transition");
        let decoded =
            Data::from_bytes(&bytes, SourceLimits::default()).expect("decode empty transition");
        assert_eq!(decoded, data);
        let inverse = decoded.inverse();
        assert_eq!(inverse.before.order.len(), 1);
        assert!(inverse.after.order.is_empty());
        assert_eq!(inverse.inverse(), decoded);
        assert_eq!(inverse.to_bytes().expect("encode inverse"), {
            let expected = decoded.inverse();
            expected.to_bytes().expect("encode expected inverse")
        });
    }

    #[test]
    fn empty_catalog_requires_complete_added_resource_presence() {
        let mut data = empty_transition();
        let mut changes = data.changes.to_vec();
        changes[0].after = None;
        data.changes = changes.into();
        assert!(data.to_bytes().is_err());

        let mut data = empty_transition();
        let mut changes = data.changes.to_vec();
        changes[0].before = Some(image(b"<unexpected/>", None));
        data.changes = changes.into();
        assert!(data.to_bytes().is_err());
    }

    #[test]
    fn disjoint_catalogs_are_valid_when_every_resource_transition_is_witnessed() {
        let removed = key(21);
        let replaced = key(22);
        let added = key(23);
        let content_type = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
        let relationship_type =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
        let data = Data {
            dependents: Arc::from([]),
            before: state(
                b"<p:presentation before='1'/>",
                vec![removed.clone(), replaced.clone()],
                21,
            ),
            after: state(b"<p:presentation after='1'/>", vec![added.clone()], 22),
            changes: vec![
                Change {
                    key: removed,
                    content_type: content_type.into(),
                    relationship_type: relationship_type.into(),
                    before: Some(image(b"<removed/>", None)),
                    after: None,
                },
                Change {
                    key: replaced,
                    content_type: content_type.into(),
                    relationship_type: relationship_type.into(),
                    before: Some(image(b"<replaced/>", Some(b""))),
                    after: None,
                },
                Change {
                    key: added,
                    content_type: content_type.into(),
                    relationship_type: relationship_type.into(),
                    before: None,
                    after: Some(image(b"<added/>", Some(b"<Relationships/>"))),
                },
            ]
            .into(),
        };
        let encoded = data.to_bytes().expect("encode disjoint catalogs");
        let decoded =
            Data::from_bytes(&encoded, SourceLimits::default()).expect("decode disjoint catalogs");
        assert_eq!(decoded, data);

        let mut incomplete = data.clone();
        let mut changes = incomplete.changes.to_vec();
        changes[2].after = None;
        incomplete.changes = changes.into();
        assert!(incomplete.to_bytes().is_err());
    }

    #[test]
    fn absent_and_present_empty_are_distinct() {
        let stable = key(1);
        let data = Data {
            dependents: Arc::from([]),
            before: state(b"<before/>", vec![stable.clone()], 1),
            after: state(b"<after/>", vec![stable.clone()], 2),
            changes: vec![Change {
                key: stable,
                content_type: "application/test".into(),
                relationship_type: "urn:test:slide".into(),
                before: Some(image(b"", None)),
                after: Some(image(b"", Some(b""))),
            }]
            .into(),
        };
        let decoded = Data::from_bytes(&data.to_bytes().expect("encode"), SourceLimits::default())
            .expect("decode");
        assert!(
            decoded.changes[0]
                .after
                .as_ref()
                .expect("after image")
                .relationships
                .as_ref()
                .expect("present empty relationships")
                .is_empty()
        );
        assert!(
            decoded.changes[0]
                .before
                .as_ref()
                .expect("before image")
                .relationships
                .is_none()
        );
    }

    #[test]
    fn invalid_counts_duplicate_keys_and_presence_are_rejected() {
        let mut bytes = sample().to_bytes().expect("encode");
        bytes[12..16].copy_from_slice(&u32::MAX.to_le_bytes());
        resign(&mut bytes);
        assert!(Data::from_bytes(&bytes, SourceLimits::default()).is_err());

        let mut duplicate = sample();
        duplicate.before.order = vec![key(1), key(1), key(3)].into();
        assert!(duplicate.to_bytes().is_err());

        let mut presence = sample();
        let mut changes = presence.changes.to_vec();
        changes[1].after = Some(image(b"unexpected", None));
        presence.changes = changes.into();
        assert!(presence.to_bytes().is_err());
    }

    #[test]
    fn checksum_truncation_trailing_and_oversize_are_rejected() {
        let bytes = sample().to_bytes().expect("encode");
        let mut changed = bytes.clone();
        changed[HEADER_BYTES] ^= 1;
        assert!(Data::from_bytes(&changed, SourceLimits::default()).is_err());
        assert!(Data::from_bytes(&bytes[..bytes.len() - 1], SourceLimits::default()).is_err());
        let mut trailing = bytes.clone();
        trailing.insert(trailing.len() - CHECKSUM_BYTES, 0);
        assert!(Data::from_bytes(&trailing, SourceLimits::default()).is_err());
        let oversized = vec![0; MAX_WIRE_BYTES + 1];
        assert!(Data::from_bytes(&oversized, SourceLimits::default()).is_err());
    }
}
