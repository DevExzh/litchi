//! Raw-preserving remapping for the object references embedded in Keynote
//! media graph payloads.
//!
//! This module is deliberately smaller than an archive clone transaction.  It
//! owns only the protobuf payload side of a clone: archive headers, build
//! order, metadata, UUIDs, and package-wide reachability remain with the
//! lifecycle owner.  Known reference paths are traversed lazily and unknown
//! fields are copied as opaque byte spans.  Consequently a payload can be
//! cloned without forcing the generated Prost model or normalising fields the
//! caller does not own.

use std::fmt;

use litchi_iwa_common::wire::WireView;
use litchi_iwa_common::{Error as WireError, WireLimits};
use litchi_iwa_common::{decode_varint_from_bytes, encode_varint_to_buffer};

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const LEGACY_STORAGE_MESSAGE_TYPE: u32 = 2_022;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;

const MAX_KNOWN_PATHS: usize = 32;

// These paths are the schema-owned TSP.Reference edges in the native archive
// messages.  Every path is expressed from the selected message root.  A path
// may cross a repeated message (the storage object tables); rewrite_message
// visits every occurrence while preserving source order.
const MOVIE_REFERENCE_PATHS: &[&[u32]] = &[
    &[1, 2],
    &[1, 6],
    &[1, 9],
    &[1, 10],
    &[1, 11],
    &[2],
    &[10],
    &[11],
    &[19],
];

const SHAPE_INFO_REFERENCE_PATHS: &[&[u32]] = &[
    &[1, 1, 2],
    &[1, 1, 6],
    &[1, 1, 9],
    &[1, 1, 10],
    &[1, 1, 11],
    &[1, 2],
    &[2],
    &[3],
    &[4],
];

const STORAGE_REFERENCE_PATHS: &[&[u32]] = &[
    &[2],
    &[5, 1, 2],
    &[7, 1, 2],
    &[8, 1, 2],
    &[9, 1, 2],
    &[11, 1, 2],
    &[12, 1, 2],
    &[15, 1, 2],
    &[16, 1, 2],
    &[17, 1, 2],
    &[18, 1, 2],
    &[21, 1, 2],
    &[22, 1, 2],
    &[23, 1, 2],
    &[27, 1, 2],
    &[28, 1, 2],
    &[25, 1, 2],
    &[26, 1, 2],
];

const CAPTION_INFO_REFERENCE_PATHS: &[&[u32]] = &[
    &[1, 1, 1, 2],
    &[1, 1, 1, 6],
    &[1, 1, 1, 9],
    &[1, 1, 1, 10],
    &[1, 1, 1, 11],
    &[1, 1, 2],
    &[1, 2],
    &[1, 3],
    &[1, 4],
    &[2],
];

/// The bounded result of one media payload rewrite.
#[derive(Debug, PartialEq, Eq)]
pub(super) struct ClonePayloadRewrite {
    payload: Vec<u8>,
    report: ClonePayloadReport,
}

impl ClonePayloadRewrite {
    /// Return the owned rewritten payload.
    #[must_use]
    #[cfg(test)]
    pub(super) fn payload(&self) -> &[u8] {
        &self.payload
    }

    /// Consume the rewrite and return its owned payload.
    #[must_use]
    pub(super) fn into_payload(self) -> Vec<u8> {
        self.payload
    }

    /// Return exact bounded resource observations from the rewrite.
    #[must_use]
    pub(super) const fn report(&self) -> ClonePayloadReport {
        self.report
    }
}

/// Resource observations for a payload clone.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ClonePayloadReport {
    fields: usize,
    work_bytes: usize,
    max_depth: usize,
    allocations: usize,
    references_seen: usize,
    references_rewritten: usize,
}

impl ClonePayloadReport {
    #[must_use]
    pub(super) const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub(super) const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub(super) const fn max_depth(self) -> usize {
        self.max_depth
    }

    #[must_use]
    pub(super) const fn allocations(self) -> usize {
        self.allocations
    }

    #[must_use]
    pub(super) const fn references_seen(self) -> usize {
        self.references_seen
    }

    #[must_use]
    #[cfg(test)]
    pub(super) const fn references_rewritten(self) -> usize {
        self.references_rewritten
    }
}

/// Failure raised by the private payload remapper.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum ClonePayloadError {
    /// The selected message type is outside this helper's schema-owned set.
    UnsupportedMessageType(u32),
    /// A known reference path had an invalid wire shape.
    InvalidReference,
    /// A mapped source reference advertised by the archive header was not
    /// present on the schema-owned paths we traversed.
    SourceReferenceWitnessMismatch,
    /// A remap pair was not strictly sorted or contains an invalid identity.
    InvalidRemap,
    /// A selected field could not be traversed safely.
    Wire(WireError),
    /// The lifecycle owner's shared budget rejected a planned allocation.
    Budget,
}

impl fmt::Display for ClonePayloadError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnsupportedMessageType(message_type) => {
                write!(
                    formatter,
                    "unsupported Keynote media payload type {message_type}"
                )
            },
            Self::InvalidReference => formatter.write_str("invalid Keynote media object reference"),
            Self::SourceReferenceWitnessMismatch => {
                formatter.write_str("Keynote media payload reference witness mismatch")
            },
            Self::InvalidRemap => formatter.write_str("invalid Keynote media object remap"),
            Self::Wire(error) => error.fmt(formatter),
            Self::Budget => {
                formatter.write_str("Keynote media lifecycle budget rejected allocation")
            },
        }
    }
}

impl std::error::Error for ClonePayloadError {}

impl From<WireError> for ClonePayloadError {
    fn from(error: WireError) -> Self {
        Self::Wire(error)
    }
}

/// Remap the schema-owned object references in one Keynote media payload.
///
/// `remap` must be sorted strictly by source identifier.  `source_object_refs`
/// is the selected archive message's aggregate `MessageInfo.object_references`
/// witness; it may be empty when the caller has no archive-header witness.
/// When present, every mapped witness identifier must occur on one of the
/// supported reference paths.  This catches a schema drift or an incomplete
/// clone before the candidate object is published.  The global remap may
/// contain identifiers belonging to sibling objects and therefore does not
/// require every pair to occur in this payload.
/// Remap one payload while charging every owned buffer through `charge`.
pub(super) fn remap_clone_payload_with_budget(
    source: &[u8],
    message_type: u32,
    remap: &[(u64, u64)],
    source_object_refs: &[u64],
    limits: WireLimits,
    charge: &mut dyn FnMut(usize) -> Result<(), ClonePayloadError>,
) -> Result<ClonePayloadRewrite, ClonePayloadError> {
    let paths = reference_paths(message_type)?;
    if source.len() > limits.max_input_bytes() {
        return Err(WireError::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::InputBytes,
            observed: source.len(),
            limit: limits.max_input_bytes(),
        }
        .into());
    }
    if remap.len() > limits.max_fields() {
        return Err(WireError::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::Fields,
            observed: remap.len(),
            limit: limits.max_fields(),
        }
        .into());
    }
    if source_object_refs.len() > limits.max_fields() {
        return Err(WireError::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::Fields,
            observed: source_object_refs.len(),
            limit: limits.max_fields(),
        }
        .into());
    }
    validate_remap(remap, charge)?;

    let mut context = RewriteContext::new(
        remap,
        source_object_refs,
        limits,
        usize::from(!remap.is_empty()),
        charge,
    );
    let payload = context.rewrite_message(source, paths, 0)?;
    context.validate_source_witness()?;
    if payload.len() > limits.max_output_bytes() {
        return Err(WireError::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::OutputBytes,
            observed: payload.len(),
            limit: limits.max_output_bytes(),
        }
        .into());
    }
    Ok(ClonePayloadRewrite {
        payload,
        report: context.finish(),
    })
}

#[cfg(test)]
pub(super) fn remap_clone_payload(
    source: &[u8],
    message_type: u32,
    remap: &[(u64, u64)],
    source_object_refs: &[u64],
    limits: WireLimits,
) -> Result<ClonePayloadRewrite, ClonePayloadError> {
    let mut charge = |_amount: usize| Ok::<(), ClonePayloadError>(());
    remap_clone_payload_with_budget(
        source,
        message_type,
        remap,
        source_object_refs,
        limits,
        &mut charge,
    )
}

fn reference_paths(message_type: u32) -> Result<&'static [&'static [u32]], ClonePayloadError> {
    match message_type {
        MOVIE_MESSAGE_TYPE => Ok(MOVIE_REFERENCE_PATHS),
        STORAGE_MESSAGE_TYPE | LEGACY_STORAGE_MESSAGE_TYPE => Ok(STORAGE_REFERENCE_PATHS),
        SHAPE_INFO_MESSAGE_TYPE => Ok(SHAPE_INFO_REFERENCE_PATHS),
        CAPTION_INFO_MESSAGE_TYPE => Ok(CAPTION_INFO_REFERENCE_PATHS),
        other => Err(ClonePayloadError::UnsupportedMessageType(other)),
    }
}

fn validate_remap(
    remap: &[(u64, u64)],
    charge: &mut dyn FnMut(usize) -> Result<(), ClonePayloadError>,
) -> Result<(), ClonePayloadError> {
    for &(source, target) in remap {
        if source == 0 || target == 0 {
            return Err(ClonePayloadError::InvalidRemap);
        }
    }
    if remap.windows(2).any(|pair| pair[0].0 >= pair[1].0) {
        return Err(ClonePayloadError::InvalidRemap);
    }
    let mut targets = Vec::new();
    let target_bytes = remap
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or_else(|| WireError::InvalidFormat("Keynote remap allocation overflow".to_owned()))?;
    if target_bytes != 0 {
        charge(target_bytes)?;
    }
    targets
        .try_reserve_exact(remap.len())
        .map_err(|_| WireError::Allocation {
            resource: "Keynote media object remap validation",
            amount: remap.len(),
        })?;
    targets.extend(remap.iter().map(|&(_source, target)| target));
    targets.sort_unstable();
    if targets.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(ClonePayloadError::InvalidRemap);
    }
    Ok(())
}

struct RewriteContext<'a, 'charge> {
    remap: &'a [(u64, u64)],
    source_object_refs: &'a [u64],
    seen_mapped_sources: Vec<u64>,
    charge: &'charge mut dyn FnMut(usize) -> Result<(), ClonePayloadError>,
    limits: WireLimits,
    fields: usize,
    work_bytes: usize,
    max_depth: usize,
    allocations: usize,
    references_seen: usize,
    references_rewritten: usize,
}

impl<'a, 'charge> RewriteContext<'a, 'charge> {
    fn new(
        remap: &'a [(u64, u64)],
        source_object_refs: &'a [u64],
        limits: WireLimits,
        allocations: usize,
        charge: &'charge mut dyn FnMut(usize) -> Result<(), ClonePayloadError>,
    ) -> Self {
        Self {
            remap,
            source_object_refs,
            seen_mapped_sources: Vec::new(),
            charge,
            limits,
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            allocations,
            references_seen: 0,
            references_rewritten: 0,
        }
    }

    fn rewrite_message(
        &mut self,
        source: &[u8],
        paths: &[&[u32]],
        depth: usize,
    ) -> Result<Vec<u8>, ClonePayloadError> {
        if depth > self.limits.max_nesting() {
            return Err(WireError::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Nesting,
                observed: depth,
                limit: self.limits.max_nesting(),
            }
            .into());
        }
        if source.len() > self.limits.max_output_bytes() {
            return Err(WireError::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::OutputBytes,
                observed: source.len(),
                limit: self.limits.max_output_bytes(),
            }
            .into());
        }
        self.max_depth = self.max_depth.max(depth);
        self.charge_work(source.len())?;
        let view = WireView::parse_with_limits(source, self.limits)?;
        self.allocations = self.allocations.checked_add(1).ok_or_else(|| {
            WireError::InvalidFormat("Keynote payload allocation count overflow".to_owned())
        })?;
        self.fields = self.fields.checked_add(view.len()).ok_or_else(|| {
            WireError::InvalidFormat("Keynote payload field count overflow".to_owned())
        })?;
        if self.fields > self.limits.max_fields() {
            return Err(WireError::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed: self.fields,
                limit: self.limits.max_fields(),
            }
            .into());
        }

        let mut output = Vec::new();
        if !source.is_empty() {
            (self.charge)(source.len())?;
        }
        output
            .try_reserve_exact(source.len())
            .map_err(|_| WireError::Allocation {
                resource: "Keynote media payload output",
                amount: source.len(),
            })?;
        self.allocations = self.allocations.checked_add(1).ok_or_else(|| {
            WireError::InvalidFormat("Keynote payload allocation count overflow".to_owned())
        })?;

        for field in view.fields() {
            let mut direct = false;
            let mut child_paths: [&[u32]; MAX_KNOWN_PATHS] = [&[]; MAX_KNOWN_PATHS];
            let mut child_count = 0usize;
            for path in paths {
                if path.first().copied() != Some(field.number()) {
                    continue;
                }
                if path.len() == 1 {
                    if direct {
                        return Err(ClonePayloadError::InvalidReference);
                    }
                    direct = true;
                } else {
                    if direct || child_count == child_paths.len() {
                        return Err(ClonePayloadError::InvalidReference);
                    }
                    child_paths[child_count] = &path[1..];
                    child_count += 1;
                }
            }

            if !direct && child_count == 0 {
                append_bytes(&mut output, field.raw(), self.limits, self.charge)?;
                continue;
            }
            if field.wire_type() != 2 {
                return Err(ClonePayloadError::InvalidReference);
            }
            field.validate_canonical_framing()?;

            let original_nested = field.payload();
            let replacement = if direct {
                self.rewrite_reference(original_nested)?
            } else {
                self.rewrite_message(original_nested, &child_paths[..child_count], depth + 1)?
            };
            if replacement == original_nested {
                append_bytes(&mut output, field.raw(), self.limits, self.charge)?;
                continue;
            }

            append_bytes(&mut output, field.key(), self.limits, self.charge)?;
            let mut length_buffer = [0u8; litchi_iwa_common::varint::MAX_BYTES];
            let length = u64::try_from(replacement.len()).map_err(|_| {
                WireError::InvalidFormat("Keynote payload length exceeds u64".to_owned())
            })?;
            let encoded_length = encode_varint_to_buffer(length, &mut length_buffer);
            append_bytes(&mut output, encoded_length, self.limits, self.charge)?;
            append_bytes(&mut output, &replacement, self.limits, self.charge)?;
        }
        if output.len() > self.limits.max_output_bytes() {
            return Err(WireError::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::OutputBytes,
                observed: output.len(),
                limit: self.limits.max_output_bytes(),
            }
            .into());
        }
        self.charge_work(output.len())?;
        Ok(output)
    }

    fn rewrite_reference(&mut self, source: &[u8]) -> Result<Vec<u8>, ClonePayloadError> {
        self.references_seen = self.references_seen.checked_add(1).ok_or_else(|| {
            WireError::InvalidFormat("Keynote reference count overflow".to_owned())
        })?;
        let view = WireView::parse_with_limits(source, self.limits)?;
        self.allocations = self.allocations.checked_add(1).ok_or_else(|| {
            WireError::InvalidFormat("Keynote payload allocation count overflow".to_owned())
        })?;
        self.fields = self.fields.checked_add(view.len()).ok_or_else(|| {
            WireError::InvalidFormat("Keynote payload field count overflow".to_owned())
        })?;
        if self.fields > self.limits.max_fields() {
            return Err(WireError::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed: self.fields,
                limit: self.limits.max_fields(),
            }
            .into());
        }
        self.charge_work(source.len())?;

        let mut identifier = None;
        let mut identifier_field = None;
        for field in view.fields() {
            if field.number() != 1 {
                continue;
            }
            if identifier.is_some() || field.wire_type() != 0 {
                return Err(ClonePayloadError::InvalidReference);
            }
            field.validate_canonical_key()?;
            let (value, width) = decode_varint_from_bytes(field.payload())
                .map_err(|_| ClonePayloadError::InvalidReference)?;
            if width != field.payload().len()
                || litchi_iwa_common::varint::encoded_len(value) != width
            {
                return Err(ClonePayloadError::InvalidReference);
            }
            identifier = Some(value);
            identifier_field = Some(field);
        }
        let identifier = identifier.ok_or(ClonePayloadError::InvalidReference)?;
        let replacement = self.lookup(identifier);
        let Some(target) = replacement else {
            return clone_bytes(source, self.limits, &mut self.allocations, self.charge);
        };
        (self.charge)(size_of::<u64>())?;
        self.seen_mapped_sources
            .try_reserve_exact(1)
            .map_err(|_| WireError::Allocation {
                resource: "Keynote media reference witness",
                amount: self.seen_mapped_sources.len() + 1,
            })?;
        self.allocations = self.allocations.checked_add(1).ok_or_else(|| {
            WireError::InvalidFormat("Keynote payload allocation count overflow".to_owned())
        })?;
        self.seen_mapped_sources.push(identifier);
        self.references_rewritten = self.references_rewritten.checked_add(1).ok_or_else(|| {
            WireError::InvalidFormat("Keynote reference count overflow".to_owned())
        })?;
        if target == identifier {
            return clone_bytes(source, self.limits, &mut self.allocations, self.charge);
        }
        if source.len() > self.limits.max_output_bytes() {
            return Err(WireError::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::OutputBytes,
                observed: source.len(),
                limit: self.limits.max_output_bytes(),
            }
            .into());
        }

        let identifier_field = identifier_field.ok_or(ClonePayloadError::InvalidReference)?;
        let mut output = Vec::new();
        if !source.is_empty() {
            (self.charge)(source.len())?;
        }
        output
            .try_reserve_exact(source.len())
            .map_err(|_| WireError::Allocation {
                resource: "Keynote media reference output",
                amount: source.len(),
            })?;
        self.allocations = self.allocations.checked_add(1).ok_or_else(|| {
            WireError::InvalidFormat("Keynote payload allocation count overflow".to_owned())
        })?;
        let mut value_buffer = [0u8; litchi_iwa_common::varint::MAX_BYTES];
        let encoded = encode_varint_to_buffer(target, &mut value_buffer);
        for field in view.fields() {
            if field.number() != 1 {
                append_bytes(&mut output, field.raw(), self.limits, self.charge)?;
            } else {
                append_bytes(
                    &mut output,
                    identifier_field.key(),
                    self.limits,
                    self.charge,
                )?;
                append_bytes(&mut output, encoded, self.limits, self.charge)?;
            }
        }
        if output.len() > self.limits.max_output_bytes() {
            return Err(WireError::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::OutputBytes,
                observed: output.len(),
                limit: self.limits.max_output_bytes(),
            }
            .into());
        }
        self.charge_work(output.len())?;
        Ok(output)
    }

    fn lookup(&self, identifier: u64) -> Option<u64> {
        self.remap
            .binary_search_by_key(&identifier, |&(source, _target)| source)
            .ok()
            .map(|index| self.remap[index].1)
    }

    fn validate_source_witness(&mut self) -> Result<(), ClonePayloadError> {
        if self.source_object_refs.is_empty() {
            return Ok(());
        }
        self.seen_mapped_sources.sort_unstable();
        self.seen_mapped_sources.dedup();
        for &source in self.source_object_refs {
            if self.lookup(source).is_some()
                && self.seen_mapped_sources.binary_search(&source).is_err()
            {
                return Err(ClonePayloadError::SourceReferenceWitnessMismatch);
            }
        }
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), ClonePayloadError> {
        self.work_bytes = self
            .work_bytes
            .checked_add(amount)
            .ok_or_else(|| WireError::InvalidFormat("Keynote payload work overflow".to_owned()))?;
        if self.work_bytes > self.limits.max_rewrite_work() {
            return Err(WireError::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                observed: self.work_bytes,
                limit: self.limits.max_rewrite_work(),
            }
            .into());
        }
        Ok(())
    }

    const fn finish(&self) -> ClonePayloadReport {
        ClonePayloadReport {
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: self.allocations,
            references_seen: self.references_seen,
            references_rewritten: self.references_rewritten,
        }
    }
}

fn append_bytes(
    output: &mut Vec<u8>,
    bytes: &[u8],
    limits: WireLimits,
    charge: &mut dyn FnMut(usize) -> Result<(), ClonePayloadError>,
) -> Result<(), ClonePayloadError> {
    let required = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| WireError::InvalidFormat("Keynote payload output overflow".to_owned()))?;
    if required > limits.max_output_bytes() {
        return Err(WireError::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::OutputBytes,
            observed: required,
            limit: limits.max_output_bytes(),
        }
        .into());
    }
    let available = output.capacity().saturating_sub(output.len());
    if bytes.len() > available {
        charge(bytes.len())?;
        output
            .try_reserve_exact(bytes.len())
            .map_err(|_| WireError::Allocation {
                resource: "Keynote media payload output",
                amount: bytes.len(),
            })?;
    }
    output.extend_from_slice(bytes);
    Ok(())
}

fn clone_bytes(
    source: &[u8],
    limits: WireLimits,
    allocations: &mut usize,
    charge: &mut dyn FnMut(usize) -> Result<(), ClonePayloadError>,
) -> Result<Vec<u8>, ClonePayloadError> {
    if source.len() > limits.max_output_bytes() {
        return Err(WireError::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::OutputBytes,
            observed: source.len(),
            limit: limits.max_output_bytes(),
        }
        .into());
    }
    let mut output = Vec::new();
    if !source.is_empty() {
        charge(source.len())?;
    }
    output
        .try_reserve_exact(source.len())
        .map_err(|_| WireError::Allocation {
            resource: "Keynote media payload clone",
            amount: source.len(),
        })?;
    output.extend_from_slice(source);
    *allocations = allocations.checked_add(1).ok_or_else(|| {
        WireError::InvalidFormat("Keynote payload allocation count overflow".to_owned())
    })?;
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_common::encode_varint_into;

    fn field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        encode_varint_into(&mut output, u64::from(number) << 3 | 2);
        encode_varint_into(&mut output, payload.len() as u64);
        output.extend_from_slice(payload);
        output
    }

    fn reference(identifier: u64) -> Vec<u8> {
        let mut output = Vec::new();
        encode_varint_into(&mut output, 8);
        encode_varint_into(&mut output, identifier);
        output
    }

    fn reference_identifier(payload: &[u8]) -> u64 {
        let view = WireView::parse(payload).expect("reference");
        let field = view
            .fields()
            .find(|field| field.number() == 1)
            .expect("identifier");
        decode_varint_from_bytes(field.payload())
            .expect("identifier varint")
            .0
    }

    fn movie_payload() -> Vec<u8> {
        let drawable = [field(10, &reference(10)), field(11, &reference(11))].concat();
        [
            field(1, &drawable),
            field(2, &reference(20)),
            field(10, &reference(30)),
            field(99, &[0x80, 0x06, 0x01]),
        ]
        .concat()
    }

    fn shape_info_payload() -> (Vec<u8>, Vec<u8>, Vec<u8>) {
        let pathsource = field(3, &[0x08, 0x0a]);
        let head_line_end = field(4, &[0x08, 0x14]);
        let drawable = field(2, &reference(10));
        let shape = [
            field(1, &drawable),
            field(2, &reference(50)),
            pathsource.clone(),
            head_line_end.clone(),
        ]
        .concat();
        let payload = [
            field(1, &shape),
            field(2, &reference(20)),
            field(3, &reference(30)),
            field(4, &reference(40)),
        ]
        .concat();
        (payload, pathsource, head_line_end)
    }

    fn limits(source: &[u8]) -> WireLimits {
        WireLimits::default()
            .with_input_bytes(source.len())
            .expect("input")
            .with_output_bytes(source.len() + 64)
            .expect("output")
            .with_fields(128)
            .expect("fields")
            .with_nesting(16)
            .expect("nesting")
            .with_rewrite_work(4096)
            .expect("work")
    }

    #[test]
    fn movie_paths_remap_known_edges_and_keep_unknown_bytes() {
        let source = movie_payload();
        let remap = [(10, 110), (20, 120), (30, 130)];
        let rewrite = remap_clone_payload(
            &source,
            MOVIE_MESSAGE_TYPE,
            &remap,
            &[10, 11, 20, 30],
            limits(&source),
        )
        .expect("movie remap");
        assert_eq!(rewrite.report.references_seen(), 4);
        assert_eq!(rewrite.report.references_rewritten(), 3);
        assert!(rewrite.payload().ends_with(&[0x80, 0x06, 0x01]));
        assert!(
            rewrite
                .payload()
                .windows(4)
                .any(|window| window == [0x52, 0x02, 0x08, 0x6e])
        );
    }

    #[test]
    fn shape_pathsource_and_line_end_are_opaque_while_private_refs_remap() {
        let (source, pathsource, head_line_end) = shape_info_payload();
        let rewrite = remap_clone_payload(
            &source,
            SHAPE_INFO_MESSAGE_TYPE,
            &[(10, 110), (20, 120), (30, 130), (40, 140), (50, 150)],
            &[10, 20, 30, 40, 50],
            limits(&source),
        )
        .expect("shape-info remap");
        let root = WireView::parse(rewrite.payload()).expect("shape-info root");
        let shape = root
            .fields()
            .find(|field| field.number() == 1)
            .expect("shape archive")
            .payload();
        let shape = WireView::parse(shape).expect("shape archive");
        assert!(
            shape
                .fields()
                .any(|field| field.number() == 3 && field.raw() == pathsource.as_slice())
        );
        assert!(
            shape
                .fields()
                .any(|field| field.number() == 4 && field.raw() == head_line_end.as_slice())
        );
        assert_eq!(
            shape
                .fields()
                .find(|field| field.number() == 2)
                .map(|field| reference_identifier(field.payload())),
            Some(150)
        );
        let drawable = shape
            .fields()
            .find(|field| field.number() == 1)
            .expect("drawable")
            .payload();
        let drawable = WireView::parse(drawable).expect("drawable");
        assert_eq!(
            drawable
                .fields()
                .find(|field| field.number() == 2)
                .map(|field| reference_identifier(field.payload())),
            Some(110)
        );
        for (number, expected) in [(2, 120), (3, 130), (4, 140)] {
            assert_eq!(
                root.fields()
                    .find(|field| field.number() == number)
                    .map(|field| reference_identifier(field.payload())),
                Some(expected)
            );
        }
    }

    #[test]
    fn mapped_header_witness_must_be_present_on_owned_paths() {
        let source = movie_payload();
        let error = remap_clone_payload(
            &source,
            MOVIE_MESSAGE_TYPE,
            &[(999, 1000)],
            &[999],
            limits(&source),
        )
        .expect_err("missing mapped witness");
        assert_eq!(error, ClonePayloadError::SourceReferenceWitnessMismatch);
    }

    #[test]
    fn invalid_remap_order_is_rejected_before_wire_work() {
        let source = movie_payload();
        let error = remap_clone_payload(
            &source,
            MOVIE_MESSAGE_TYPE,
            &[(20, 120), (10, 110)],
            &[],
            limits(&source),
        )
        .expect_err("unsorted remap");
        assert_eq!(error, ClonePayloadError::InvalidRemap);
    }

    #[test]
    fn duplicate_remap_targets_are_rejected_even_when_non_adjacent() {
        let source = movie_payload();
        let error = remap_clone_payload(
            &source,
            MOVIE_MESSAGE_TYPE,
            &[(10, 110), (20, 120), (30, 110)],
            &[],
            limits(&source),
        )
        .expect_err("duplicate target");
        assert_eq!(error, ClonePayloadError::InvalidRemap);
    }

    #[test]
    fn shared_budget_is_charged_before_output_reserve() {
        let source = movie_payload();
        let mut charged = Vec::new();
        let mut charge = |amount: usize| {
            charged.push(amount);
            Err(ClonePayloadError::Budget)
        };
        let error = remap_clone_payload_with_budget(
            &source,
            MOVIE_MESSAGE_TYPE,
            &[],
            &[],
            limits(&source),
            &mut charge,
        )
        .expect_err("shared budget");
        assert_eq!(error, ClonePayloadError::Budget);
        assert_eq!(charged, vec![source.len()]);
    }

    #[test]
    fn unsupported_message_type_does_not_perform_opaque_clone() {
        let error = remap_clone_payload(&[], 9999, &[], &[], WireLimits::default())
            .expect_err("unsupported type");
        assert_eq!(error, ClonePayloadError::UnsupportedMessageType(9999));
    }

    #[test]
    fn exact_noop_preserves_payload_bytes() {
        let source = movie_payload();
        let rewrite = remap_clone_payload(
            &source,
            MOVIE_MESSAGE_TYPE,
            &[(10, 110), (20, 120), (30, 130)],
            &[10, 11, 20, 30],
            limits(&source),
        )
        .expect("movie remap");
        let again = remap_clone_payload(
            rewrite.payload(),
            MOVIE_MESSAGE_TYPE,
            &[(110, 110)],
            &[],
            limits(rewrite.payload()),
        )
        .expect("noop remap");
        assert_eq!(again.payload(), rewrite.payload());
    }
}
