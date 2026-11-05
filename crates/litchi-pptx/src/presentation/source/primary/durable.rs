//! Versioned, bounded transport for primary source patches.
//!
//! Version 2 carries the exact presentation-root and stable catalog order
//! preimages.  A text patch authored after a reorder therefore cannot be
//! applied to the original logical order, while an inverse patch remains
//! replayable after reopening the reordered ZIP: the digest canonicalizes the
//! root's order entries while retaining every other root byte, and the exact
//! root/order bytes provide the state guard.

use super::super::{
    MAX_SOURCE_BACKED_SLIDE_BATCH, PartBinding, RelationshipBinding, SourcePayload,
};
use super::{Package, Patch, PatchData, RootPayload, allocation, limit};
use crate::presentation::order::{
    StableSlideKey, canonical_source_order_xml, reorder_stable_bindings, stable_slide_key,
    validate_source_order_root,
};
use crate::{Error, Result};
use sha2::{Digest, Sha256};
use std::sync::{Arc, OnceLock};

const REMOVAL_MAGIC: &[u8; 8] = b"LPSX0004";
pub(super) const REMOVAL_MARKER_BYTES: usize = REMOVAL_MAGIC.len();
const MAGIC: &[u8; 8] = b"LPSX0002";
const HEADER: usize = 8 + 32 + 4 + 4 + 4 + 8 + 8;
const KEY_FIXED: usize = 4 + 4 + 4;

#[derive(Clone)]
pub(super) struct Data {
    pub(super) root: [u8; 32],
    pub(super) before_root: Arc<Vec<u8>>,
    pub(super) after_root: Arc<Vec<u8>>,
    pub(super) before_order: Arc<[StableSlideKey]>,
    pub(super) after_order: Arc<[StableSlideKey]>,
    pub(super) changes: Arc<[Change]>,
}

#[derive(Clone)]
pub(super) struct Change {
    key: StableSlideKey,
    pub(super) before: Arc<Vec<u8>>,
    pub(super) after: Arc<Vec<u8>>,
}

impl Data {
    pub(super) fn inverse(&self) -> Self {
        Self {
            root: self.root,
            before_root: Arc::clone(&self.after_root),
            after_root: Arc::clone(&self.before_root),
            before_order: Arc::clone(&self.after_order),
            after_order: Arc::clone(&self.before_order),
            changes: self
                .changes
                .iter()
                .map(|change| Change {
                    key: change.key.clone(),
                    before: Arc::clone(&change.after),
                    after: Arc::clone(&change.before),
                })
                .collect(),
        }
    }
}

impl Patch {
    /// Encode a bounded exact source patch for replay after reopening its
    /// source.  The stable catalog keys are physical identities, so all wire
    /// positions continue to refer to the same slide after a reorder.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if !self.shape_removals {
            if let PatchData::Topology(data) = &self.data {
                return data.to_bytes();
            }
        }
        let length = self.encoded_len()?;
        let mut output = Vec::new();
        output.try_reserve_exact(length).map_err(allocation)?;
        if self.shape_removals {
            output.extend_from_slice(REMOVAL_MAGIC);
        }
        match &self.data {
            PatchData::Local {
                source,
                before_root,
                before_order,
                after_root,
                after_order,
                changes,
            } => {
                source.check()?;
                write_header(
                    &mut output,
                    root_digest(source)?,
                    before_order.len(),
                    after_order.len(),
                    changes.len(),
                    before_root.as_bytes().len(),
                    after_root.as_bytes().len(),
                )?;
                write_order(&mut output, before_order)?;
                write_order(&mut output, after_order)?;
                output.extend_from_slice(before_root.as_bytes());
                output.extend_from_slice(after_root.as_bytes());
                for change in changes.iter() {
                    let key = super::source_key_for_position(source, change.before.position)?;
                    put_key(&mut output, &key)?;
                    write_payload(&mut output, change.before.xml.as_bytes());
                    write_payload(&mut output, change.after.xml.as_bytes());
                }
                source.check()?;
            },
            PatchData::Durable(data) => {
                write_header(
                    &mut output,
                    data.root,
                    data.before_order.len(),
                    data.after_order.len(),
                    data.changes.len(),
                    data.before_root.len(),
                    data.after_root.len(),
                )?;
                write_order(&mut output, &data.before_order)?;
                write_order(&mut output, &data.after_order)?;
                output.extend_from_slice(&data.before_root);
                output.extend_from_slice(&data.after_root);
                for change in data.changes.iter() {
                    put_key(&mut output, &change.key)?;
                    write_payload(&mut output, &change.before);
                    write_payload(&mut output, &change.after);
                }
            },
            PatchData::Topology(data) => output.extend_from_slice(&data.to_bytes()?),
        }
        if output.len() != length {
            return Err(invalid("wire length accounting mismatch"));
        }
        Ok(output)
    }

    /// Decode a bounded source patch. Version 2 retains text/permutation
    /// semantics; version 3 additionally carries guarded topology changes.
    /// Version 4 wraps either format and admits the managed slide name/hidden
    /// grammar as well as reversible top-level shape removal. Earlier versions
    /// retain their narrower addition grammar.
    /// Framing, complete root/order
    /// state, stable-key uniqueness, and XML pairs are validated before any
    /// retained payload is allocated.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() > maximum() {
            return Err(limit(
                crate::error::CodecLimitKind::PatchBytes,
                bytes.len(),
                maximum(),
            ));
        }
        let (shape_removals, bytes) = bytes
            .strip_prefix(REMOVAL_MAGIC)
            .map_or((false, bytes), |body| (true, body));
        if bytes.starts_with(b"LPSX0003")
            || bytes.starts_with(super::topology_wire::DEPENDENT_MAGIC)
        {
            return Ok(Self {
                shape_removals,
                data: PatchData::Topology(Arc::new(super::topology::Patch::from_bytes(bytes)?)),
                encoded_len_cache: Arc::new(OnceLock::new()),
            });
        }
        if bytes.len() > maximum() {
            return Err(limit(
                crate::error::CodecLimitKind::PatchBytes,
                bytes.len(),
                maximum(),
            ));
        }
        let mut input = Input { bytes, offset: 0 };
        if input.take(MAGIC.len())? != MAGIC {
            return Err(invalid("unsupported wire version"));
        }
        let root = input.digest()?;
        let before_count = input.count()?;
        let after_count = input.count()?;
        let change_count = input.count()?;
        if before_count > crate::parts::MAX_SLIDES || after_count > crate::parts::MAX_SLIDES {
            return Err(limit(
                crate::error::CodecLimitKind::Slides,
                before_count.max(after_count),
                crate::parts::MAX_SLIDES,
            ));
        }
        let catalog_keys = before_count
            .checked_add(after_count)
            .ok_or_else(|| invalid("root order count overflows"))?;
        if change_count > MAX_SOURCE_BACKED_SLIDE_BATCH {
            return Err(limit(
                crate::error::CodecLimitKind::Slides,
                change_count,
                MAX_SOURCE_BACKED_SLIDE_BATCH,
            ));
        }
        if change_count > bytes.len() / (KEY_FIXED + 8 + 8) {
            return Err(invalid("change framing exceeds input"));
        }
        let before_root_len = input.length()?;
        let after_root_len = input.length()?;
        if before_root_len > maximum() || after_root_len > maximum() {
            return Err(limit(
                crate::error::CodecLimitKind::PatchBytes,
                before_root_len.max(after_root_len),
                maximum(),
            ));
        }
        let minimum = catalog_keys
            .checked_mul(KEY_FIXED)
            .and_then(|size| size.checked_add(before_root_len))
            .and_then(|size| size.checked_add(after_root_len))
            .and_then(|size| size.checked_add(change_count.checked_mul(KEY_FIXED + 16)?))
            .ok_or_else(|| invalid("framing size overflows"))?;
        if minimum > bytes.len().saturating_sub(input.offset) {
            return Err(invalid("root order framing exceeds input"));
        }
        let before_order = read_order(&mut input, before_count)?;
        let after_order = read_order(&mut input, after_count)?;
        validate_order_sets(&before_order, &after_order)?;
        let before_root_bytes = input.take(before_root_len)?;
        let after_root_bytes = input.take(after_root_len)?;
        validate_root_delta(
            before_root_bytes,
            after_root_bytes,
            &before_order,
            &after_order,
        )?;
        let mut changes = Vec::new();
        changes
            .try_reserve_exact(change_count)
            .map_err(allocation)?;
        let mut seen = std::collections::HashSet::new();
        seen.try_reserve(change_count).map_err(allocation)?;
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(change_count)
            .map_err(allocation)?;
        for _ in 0..change_count {
            let key = read_key(&mut input)?;
            if !before_order.iter().any(|candidate| candidate == &key) {
                return Err(invalid("change references an unknown stable slide key"));
            }
            if !seen.insert(key.clone()) {
                return Err(invalid("durable patch repeats a stable slide key"));
            }
            let before = input.payload()?;
            let after = input.payload()?;
            if before == after {
                return Err(invalid("redundant no-op delta"));
            }
            validate_pair(before, after, None, None, shape_removals)?;
            entries.push((key, before, after));
        }
        if input.offset != bytes.len() {
            return Err(invalid("trailing bytes"));
        }
        let before_root = copy(before_root_bytes)?;
        let after_root = copy(after_root_bytes)?;
        for (key, before, after) in entries {
            changes.push(Change {
                key,
                before: copy(before)?,
                after: copy(after)?,
            });
        }
        Ok(Self {
            shape_removals,
            data: PatchData::Durable(Data {
                root,
                before_root,
                after_root,
                before_order: before_order.into(),
                after_order: after_order.into(),
                changes: changes.into(),
            }),
            encoded_len_cache: Arc::new(OnceLock::new()),
        })
    }
}

pub(super) fn apply(package: &Package, data: &Data, shape_removals: bool) -> Result<Package> {
    package.check()?;
    if root_digest(package)? != data.root
        || package.root.as_bytes() != data.before_root.as_slice()
        || package.order.as_ref() != data.before_order.as_ref()
    {
        return Err(Error::PatchConflict);
    }
    let root_changed = data.before_root.as_slice() != data.after_root.as_slice();
    let order_changed = data.before_order.as_ref() != data.after_order.as_ref();
    // A root may change without an order change only in section membership
    // (a slide moved out and back across a section boundary, record 0303).
    if (order_changed && !root_changed)
        || (root_changed
            && !order_changed
            && !order_delta_matches(
                &data.before_root,
                &data.before_order,
                &data.after_order,
                &data.after_root,
                package.owner.package.execution_context().as_ref(),
            )?)
    {
        return Err(Error::PatchConflict);
    }
    if root_changed {
        validate_source_order_root(
            &package.owner.package,
            package.owner._presentation.source_partname(),
            &data.before_root,
            &data.before_order,
            package.owner.package.execution_context().as_ref(),
        )?;
        validate_source_order_root(
            &package.owner.package,
            package.owner._presentation.source_partname(),
            &data.after_root,
            &data.after_order,
            package.owner.package.execution_context().as_ref(),
        )?;
    }

    // Validate all physical slide preimages and output limits before creating
    // a candidate package or allocating any replacement payload.
    let mut retained_bytes = if root_changed {
        data.after_root.len()
    } else {
        0
    };
    for change in data.changes.iter() {
        package.check()?;
        let physical = package.physical_position(&change.key)?;
        let current = package.snapshot(physical)?;
        if current.xml.as_bytes() != change.before.as_slice() {
            return Err(Error::PatchConflict);
        }
        if change.after.len() > current.max_output_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::OutputBytes,
                change.after.len(),
                current.max_output_bytes,
            ));
        }
        let semantic = crate::open_options::selected(package.owner.semantic_limits.as_ref());
        validate_pair_with_limits(
            &change.before,
            &change.after,
            Some(package.owner.limits),
            package.owner.package.execution_context().as_ref(),
            shape_removals,
            semantic.shape(),
            semantic.transition(),
        )?;
        retained_bytes = retained_bytes
            .checked_add(change.after.len())
            .ok_or_else(|| invalid("candidate size overflows"))?;
    }
    let changed = package.root.as_bytes() != data.after_root.as_slice()
        || package.order.as_ref() != data.after_order.as_ref()
        || !data.changes.is_empty();
    if changed && package.owner.package.has_signature_infrastructure() {
        return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
    }
    let reservation = if changed {
        package
            .owner
            .package
            .execution_context()
            .map(|context| {
                context
                    .reserve(litchi_core::Resource::Memory, retained_bytes as u64)
                    .map(Arc::new)
                    .map_err(execution_error)
            })
            .transpose()?
    } else {
        None
    };
    if root_changed
        && !order_delta_matches(
            &data.before_root,
            &data.before_order,
            &data.after_order,
            &data.after_root,
            package.owner.package.execution_context().as_ref(),
        )?
    {
        return Err(Error::PatchConflict);
    }
    let root = if let Some(reservation) = &reservation {
        RootPayload::ReservedEdit {
            bytes: Arc::clone(&data.after_root),
            _reservation: Arc::clone(reservation),
        }
    } else {
        RootPayload::Edited(Arc::clone(&data.after_root))
    };
    let mut candidate = package.with_order(root, Arc::clone(&data.after_order))?;
    for change in data.changes.iter() {
        let physical = candidate.physical_position(&change.key)?;
        let mut target = package.snapshot(physical)?;
        target.position = physical;
        target.xml = match &reservation {
            Some(reservation) => SourcePayload::ReservedEdit {
                bytes: Arc::clone(&change.after),
                reservation: Arc::clone(reservation),
            },
            None => SourcePayload::Edited(Arc::clone(&change.after)),
        };
        candidate = candidate.with_snapshot(target)?;
    }
    candidate.check()?;
    Ok(candidate)
}

pub(super) fn validate_pair(
    before: &[u8],
    after: &[u8],
    limits: Option<litchi_opc::SourceLimits>,
    context: Option<&litchi_core::ExecutionContext>,
    shape_removals: bool,
) -> Result<()> {
    validate_pair_with_limits(
        before,
        after,
        limits,
        context,
        shape_removals,
        crate::shape::Limits::DEFAULT,
        crate::transition::Limits::DEFAULT,
    )
}

pub(super) fn validate_pair_with_limits(
    before: &[u8],
    after: &[u8],
    limits: Option<litchi_opc::SourceLimits>,
    context: Option<&litchi_core::ExecutionContext>,
    shape_removals: bool,
    shape_limits: crate::shape::Limits,
    transition_limits: crate::transition::Limits,
) -> Result<()> {
    let _proof_reservation = context
        .map(|context| {
            let maximum = crate::opened::Limits::default().max_patch_bytes();
            let bytes = before
                .len()
                .checked_add(after.len())
                .and_then(|bytes| bytes.checked_mul(64))
                .and_then(|bytes| bytes.checked_add(1024))
                .ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::ProofBytes,
                        usize::MAX,
                        maximum,
                    )
                })?;
            if bytes > maximum {
                return Err(limit(
                    crate::error::CodecLimitKind::ProofBytes,
                    bytes,
                    maximum,
                ));
            }
            let amount = u64::try_from(bytes).map_err(|_| {
                limit(
                    crate::error::CodecLimitKind::ProofBytes,
                    usize::MAX,
                    maximum,
                )
            })?;
            context
                .reserve(litchi_core::Resource::Memory, amount)
                .map_err(|error| {
                    Error::Opc(match error {
                        litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                        other => litchi_opc::OpcError::Execution(other),
                    })
                })
        })
        .transpose()?;
    // A slide that encloses markup compatibility in shape-tree children
    // admits only whole top-level shape removal, proven on the source bytes
    // (record 0303).
    if shape_removals {
        if let Some(limits) = limits {
            for xml in [before, after] {
                crate::source_limits::validate_xml_limits(xml, limits)?;
            }
        }
        if crate::opened::validate_enclosed_removal_pair_with_limits(
            before,
            after,
            "apply",
            shape_limits,
            transition_limits,
        )? {
            return Ok(());
        }
    }
    for xml in [before, after] {
        if let Some(limits) = limits {
            if xml.len() as u64 > limits.max_part_bytes() {
                return Err(Error::Opc(litchi_opc::OpcError::Limit {
                    kind: litchi_opc::SourceLimitKind::PartBytes,
                    observed: xml.len() as u64,
                    maximum: limits.max_part_bytes(),
                }));
            }
            crate::source_limits::validate_xml_limits(xml, limits)?;
        }
        crate::presentation::transition::validate_source_with_limits(
            xml,
            "apply",
            shape_limits,
            transition_limits,
        )?;
    }
    // Old wire versions retain their original shape-only proof grammar.
    if !shape_removals {
        let before_skeleton = crate::opened::shape_text_skeleton_with_limits(before, shape_limits)?;
        let after_skeleton = crate::opened::shape_text_skeleton_with_limits(after, shape_limits)?;
        if before_skeleton == after_skeleton
            || validate_property_surface(
                before,
                after,
                before_skeleton,
                after_skeleton,
                context,
                false,
                shape_limits,
            )?
        {
            return Ok(());
        }
        return Err(invalid(
            "patch changes XML outside editable shape properties",
        ));
    }

    // The expanded envelope admits managed slide attributes. A matching
    // rooted skeleton proves an attribute-only edit without shape decoding.
    // Both scans validate managed values before erasing their lexical spans.
    if crate::opened::slide_properties_skeleton(before)?
        == crate::opened::slide_properties_skeleton(after)?
    {
        return Ok(());
    }
    let normalized_before = crate::opened::normalize_slide_properties(before)?;
    let normalized_after = crate::opened::normalize_slide_properties(after)?;
    let before_skeleton =
        crate::opened::shape_text_skeleton_with_limits(&normalized_before, shape_limits)?;
    let after_skeleton =
        crate::opened::shape_text_skeleton_with_limits(&normalized_after, shape_limits)?;
    if validate_property_surface(
        &normalized_before,
        &normalized_after,
        before_skeleton,
        after_skeleton,
        context,
        true,
        shape_limits,
    )? {
        return Ok(());
    }
    Err(invalid(
        "patch changes XML outside editable shape properties or top-level removal",
    ))
}

/// Apply the existing bounded text, geometry, append, and optional removal
/// proof to a pair whose managed slide attributes have already been validated
/// and, when needed, erased by the expanded proof caller.
fn validate_property_surface(
    before: &[u8],
    after: &[u8],
    before_skeleton: [u8; 32],
    after_skeleton: [u8; 32],
    context: Option<&litchi_core::ExecutionContext>,
    shape_removals: bool,
    shape_limits: crate::shape::Limits,
) -> Result<bool> {
    if before_skeleton == after_skeleton {
        return Ok(true);
    }
    // Geometry is a separate, narrowly located surface. Mask only validated
    // direct transform values, then apply the existing exact text-slot proof.
    let properties_equal = |left: &[u8], right: &[u8]| -> Result<bool> {
        let _geometry_reservation = context
            .map(|context| {
                let maximum = crate::opened::Limits::default().max_patch_bytes();
                let bytes = left
                    .len()
                    .checked_add(right.len())
                    .and_then(|size| size.checked_mul(64))
                    .and_then(|size| size.checked_add(1024))
                    .ok_or_else(|| {
                        limit(
                            crate::error::CodecLimitKind::ProofBytes,
                            usize::MAX,
                            maximum,
                        )
                    })?;
                if bytes > maximum {
                    return Err(limit(
                        crate::error::CodecLimitKind::ProofBytes,
                        bytes,
                        maximum,
                    ));
                }
                let amount = u64::try_from(bytes).map_err(|_| {
                    limit(
                        crate::error::CodecLimitKind::ProofBytes,
                        usize::MAX,
                        maximum,
                    )
                })?;
                context
                    .reserve(litchi_core::Resource::Memory, amount)
                    .map_err(|error| {
                        Error::Opc(match error {
                            litchi_core::ExecutionError::Cancelled => {
                                litchi_opc::OpcError::Cancelled
                            },
                            other => litchi_opc::OpcError::Execution(other),
                        })
                    })
            })
            .transpose()?;
        let left = crate::opened::geometry_skeleton_with_limits(left, shape_limits)?;
        let right = crate::opened::geometry_skeleton_with_limits(right, shape_limits)?;
        Ok(
            crate::opened::shape_text_skeleton_with_limits(&left, shape_limits)?
                == crate::opened::shape_text_skeleton_with_limits(&right, shape_limits)?,
        )
    };
    if properties_equal(before, after)? {
        return Ok(true);
    }
    let unchanged = if let Some(stripped_after) =
        crate::opened::strip_appended_text_boxes_with_limits(before, after, shape_limits)?
    {
        crate::opened::shape_text_skeleton_with_limits(&stripped_after, shape_limits)?
            == before_skeleton
            || properties_equal(before, &stripped_after)?
    } else if let Some(stripped_before) =
        crate::opened::strip_appended_text_boxes_with_limits(after, before, shape_limits)?
    {
        crate::opened::shape_text_skeleton_with_limits(&stripped_before, shape_limits)?
            == after_skeleton
            || properties_equal(&stripped_before, after)?
    } else {
        false
    };
    if unchanged {
        return Ok(true);
    }
    if shape_removals {
        // A reversible root deletion/restoration may compose with retained
        // text and geometry edits plus canonical appended text boxes.
        for (forward, replacements) in [(true, false), (false, false), (true, true), (false, true)]
        {
            if let Some(pair) = crate::opened::normalize_shape_removal_pair_with_limits(
                before,
                after,
                forward,
                replacements,
                shape_limits,
            )? {
                if crate::opened::shape_text_skeleton_with_limits(&pair.before, shape_limits)?
                    == crate::opened::shape_text_skeleton_with_limits(&pair.after, shape_limits)?
                    || properties_equal(&pair.before, &pair.after)?
                {
                    return Ok(true);
                }
            }
        }
    }
    Ok(false)
}

/// Digest the physical source topology and a canonicalized presentation root.
/// The canonical root is independent of slide-entry order but retains every
/// other root byte, so this identifies the source owner across a reopen of a
/// valid reordered ZIP without allowing a forged root pair to rebind a patch.
fn root_digest(package: &Package) -> Result<[u8; 32]> {
    let mut hash = Sha256::new();
    hash.update(b"litchi-pptx-source-topology-v2\0");
    relationships(&mut hash, &package.owner.package_relationships);
    part(&mut hash, &package.owner.presentation_binding);
    let mut slides = package
        .owner
        .slides
        .iter()
        .map(|slide| {
            (
                stable_slide_key(
                    slide.part_uri.clone(),
                    slide.slide_id,
                    slide.binding.slide_reference_id.clone(),
                ),
                slide.binding.clone(),
            )
        })
        .collect::<Vec<_>>();
    slides.sort_by(|left, right| {
        left.0
            .part_uri
            .as_str()
            .cmp(right.0.part_uri.as_str())
            .then(left.0.slide_id.cmp(&right.0.slide_id))
            .then(left.0.relationship_id.cmp(&right.0.relationship_id))
    });
    hash.update((slides.len() as u64).to_le_bytes());
    for (key, binding) in slides {
        stable_key_hash(&mut hash, &key);
        relationship(&mut hash, &binding.presentation_relationship);
        part(&mut hash, &binding.slide);
    }
    let canonical = canonical_source_order_xml(
        package.root.as_bytes(),
        &package.order,
        package.owner.package.execution_context().as_ref(),
    )?;
    field(&mut hash, &canonical);
    Ok(hash.finalize().into())
}

fn stable_key_hash(hash: &mut Sha256, key: &StableSlideKey) {
    field(hash, key.part_uri.as_str().as_bytes());
    hash.update(key.slide_id.to_le_bytes());
    field(hash, key.relationship_id.as_bytes());
}

fn field(hash: &mut Sha256, value: &[u8]) {
    hash.update((value.len() as u64).to_le_bytes());
    hash.update(value);
}
fn relationship(hash: &mut Sha256, value: &RelationshipBinding) {
    field(hash, value.id.as_bytes());
    field(hash, value.kind.as_bytes());
    field(hash, value.target.as_bytes());
    hash.update([u8::from(matches!(
        value.mode,
        litchi_opc::TargetMode::External
    ))]);
}
fn relationships(hash: &mut Sha256, values: &[RelationshipBinding]) {
    hash.update((values.len() as u64).to_le_bytes());
    for value in values {
        relationship(hash, value);
    }
}
fn part(hash: &mut Sha256, value: &PartBinding) {
    field(hash, value.uri.as_str().as_bytes());
    field(hash, value.content_type.as_bytes());
    relationships(hash, &value.relationships);
}

fn maximum() -> usize {
    crate::opened::Limits::default().max_patch_bytes()
}

/// Size the local source patch envelope without retaining any payload.  The
/// primary editor uses this before freezing a focused slide edit so the
/// caller's aggregate patch ceiling covers the complete before/after root,
/// order and slide framing, rather than only the new XML.
pub(super) fn local_wire_len(
    shape_removals: bool,
    before_order: &[StableSlideKey],
    after_order: &[StableSlideKey],
    before_root: usize,
    after_root: usize,
    changes: &[(StableSlideKey, usize, usize)],
) -> Result<usize> {
    let mut total = HEADER
        .checked_add(if shape_removals {
            REMOVAL_MAGIC.len()
        } else {
            0
        })
        .ok_or_else(|| invalid("wire length overflows"))?;
    total = body_len(total, before_order, after_order, before_root, after_root)?;
    for (key, before, after) in changes {
        total = entry_len(total, key, *before, *after)?;
    }
    Ok(total)
}

/// Size the durable (already detached from a source owner) patch envelope.
///
/// This is deliberately shared with the primary facade's accounting path:
/// the `Data` fields are the authoritative wire representation, while the
/// outer `Patch` adds the optional removal marker when it is enabled.
pub(super) fn data_wire_len(data: &Data) -> Result<usize> {
    let mut total = HEADER;
    total = body_len(
        total,
        &data.before_order,
        &data.after_order,
        data.before_root.len(),
        data.after_root.len(),
    )?;
    for change in data.changes.iter() {
        total = entry_len(total, &change.key, change.before.len(), change.after.len())?;
    }
    if data.changes.len() > MAX_SOURCE_BACKED_SLIDE_BATCH {
        return Err(limit(
            crate::error::CodecLimitKind::Slides,
            data.changes.len(),
            MAX_SOURCE_BACKED_SLIDE_BATCH,
        ));
    }
    Ok(total)
}

fn body_len(
    total: usize,
    before_order: &[StableSlideKey],
    after_order: &[StableSlideKey],
    before_root: usize,
    after_root: usize,
) -> Result<usize> {
    u32::try_from(before_order.len()).map_err(|_| invalid("root order exceeds wire width"))?;
    u32::try_from(after_order.len()).map_err(|_| invalid("root order exceeds wire width"))?;
    let mut total = total
        .checked_add(before_root)
        .and_then(|value| value.checked_add(after_root))
        .ok_or_else(|| invalid("wire length overflows"))?;
    for key in before_order.iter().chain(after_order) {
        total = total
            .checked_add(key_len(key)?)
            .ok_or_else(|| invalid("wire length overflows"))?;
    }
    if total > maximum() {
        return Err(limit(
            crate::error::CodecLimitKind::PatchBytes,
            total,
            maximum(),
        ));
    }
    Ok(total)
}

fn entry_len(total: usize, key: &StableSlideKey, before: usize, after: usize) -> Result<usize> {
    let length = total
        .checked_add(key_len(key)?)
        .and_then(|value| value.checked_add(8 + 8))
        .and_then(|value| value.checked_add(before))
        .and_then(|value| value.checked_add(after))
        .ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::PatchBytes,
                usize::MAX,
                maximum(),
            )
        })?;
    if length > maximum() {
        return Err(limit(
            crate::error::CodecLimitKind::PatchBytes,
            length,
            maximum(),
        ));
    }
    Ok(length)
}

fn key_len(key: &StableSlideKey) -> Result<usize> {
    let uri = key.part_uri.as_str().len();
    let relationship = key.relationship_id.len();
    u32::try_from(uri).map_err(|_| invalid("part URI exceeds wire width"))?;
    u32::try_from(relationship).map_err(|_| invalid("relationship ID exceeds wire width"))?;
    KEY_FIXED
        .checked_add(uri)
        .and_then(|value| value.checked_add(relationship))
        .ok_or_else(|| invalid("stable key length overflows"))
}

fn write_header(
    output: &mut Vec<u8>,
    root: [u8; 32],
    before_count: usize,
    after_count: usize,
    change_count: usize,
    before_root_len: usize,
    after_root_len: usize,
) -> Result<()> {
    output.extend_from_slice(MAGIC);
    output.extend_from_slice(&root);
    put_u32(output, before_count)?;
    put_u32(output, after_count)?;
    put_u32(output, change_count)?;
    output.extend_from_slice(
        &u64::try_from(before_root_len)
            .map_err(|_| invalid("root payload exceeds wire width"))?
            .to_le_bytes(),
    );
    output.extend_from_slice(
        &u64::try_from(after_root_len)
            .map_err(|_| invalid("root payload exceeds wire width"))?
            .to_le_bytes(),
    );
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

fn write_order(output: &mut Vec<u8>, order: &[StableSlideKey]) -> Result<()> {
    for key in order {
        put_key(output, key)?;
    }
    Ok(())
}

fn put_key(output: &mut Vec<u8>, key: &StableSlideKey) -> Result<()> {
    let uri = key.part_uri.as_str().as_bytes();
    let relationship = key.relationship_id.as_bytes();
    put_u32(output, uri.len())?;
    output.extend_from_slice(uri);
    output.extend_from_slice(&key.slide_id.to_le_bytes());
    put_u32(output, relationship.len())?;
    output.extend_from_slice(relationship);
    Ok(())
}

fn write_payload(output: &mut Vec<u8>, bytes: &[u8]) {
    output.extend_from_slice(&(bytes.len() as u64).to_le_bytes());
    output.extend_from_slice(bytes);
}

fn validate_order_sets(before: &[StableSlideKey], after: &[StableSlideKey]) -> Result<()> {
    if before.len() != after.len() {
        return Err(invalid("root order lengths differ"));
    }
    let mut before_set = std::collections::HashSet::new();
    let mut after_set = std::collections::HashSet::new();
    before_set.try_reserve(before.len()).map_err(allocation)?;
    after_set.try_reserve(after.len()).map_err(allocation)?;
    for key in before {
        if !before_set.insert(key) {
            return Err(invalid("root order repeats a stable key"));
        }
    }
    for key in after {
        if !after_set.insert(key) {
            return Err(invalid("root order repeats a stable key"));
        }
    }
    if before_set != after_set {
        return Err(invalid("root order changes the physical slide catalog"));
    }
    Ok(())
}

fn validate_root_delta(
    before_root: &[u8],
    after_root: &[u8],
    before_order: &[StableSlideKey],
    after_order: &[StableSlideKey],
) -> Result<()> {
    let root_changed = before_root != after_root;
    let order_changed = before_order != after_order;
    if order_changed && !root_changed {
        return Err(invalid("root and stable order changed inconsistently"));
    }
    // A root change without an order change may move section entries only.
    if root_changed
        && !order_delta_matches(before_root, before_order, after_order, after_root, None)?
    {
        if !order_changed {
            return Err(invalid("root and stable order changed inconsistently"));
        }
        return Err(invalid("root delta is not a native order permutation"));
    }
    Ok(())
}

/// Whether `after_root` is `before_root` with its slide catalog permuted to
/// `after_order`. A sectioned root may also move section entries; their
/// membership must be one the order change keeps (record 0303).
fn order_delta_matches(
    before_root: &[u8],
    before_order: &[StableSlideKey],
    after_order: &[StableSlideKey],
    after_root: &[u8],
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<bool> {
    let expected = reorder_stable_bindings(before_root, before_order, after_order, context)?;
    if expected.as_slice() == after_root {
        return Ok(true);
    }
    if !crate::opened::sections::present(before_root) {
        return Ok(false);
    }
    let (Some((expected, _)), Some((actual, after_members)), Some((_, before_members))) = (
        crate::opened::sections::strip(&expected)?,
        crate::opened::sections::strip(after_root)?,
        crate::opened::sections::strip(before_root)?,
    ) else {
        return Ok(false);
    };
    let before_ids: Vec<u32> = before_order.iter().map(|key| key.slide_id).collect();
    let after_ids: Vec<u32> = after_order.iter().map(|key| key.slide_id).collect();
    crate::opened::sections::validate_membership(
        &before_members,
        &after_members,
        &before_ids,
        &after_ids,
    )?;
    Ok(expected == actual)
}

fn copy(bytes: &[u8]) -> Result<Arc<Vec<u8>>> {
    let mut output = Vec::new();
    output.try_reserve_exact(bytes.len()).map_err(allocation)?;
    output.extend_from_slice(bytes);
    Ok(Arc::new(output))
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(format!("source-backed durable patch: {message}"))
}
fn execution_error(error: litchi_core::ExecutionError) -> Error {
    Error::Opc(match error {
        litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
        error => litchi_opc::OpcError::Execution(error),
    })
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
            .ok_or_else(|| invalid("offset overflow"))?;
        let bytes = self
            .bytes
            .get(self.offset..end)
            .ok_or_else(|| invalid("truncated encoding"))?;
        self.offset = end;
        Ok(bytes)
    }
    fn count(&mut self) -> Result<usize> {
        let bytes: [u8; 4] = self
            .take(4)?
            .try_into()
            .map_err(|_| invalid("integer width"))?;
        usize::try_from(u32::from_le_bytes(bytes)).map_err(|_| invalid("integer overflow"))
    }
    fn length(&mut self) -> Result<usize> {
        let bytes: [u8; 8] = self
            .take(8)?
            .try_into()
            .map_err(|_| invalid("length width"))?;
        usize::try_from(u64::from_le_bytes(bytes)).map_err(|_| invalid("length overflow"))
    }
    fn digest(&mut self) -> Result<[u8; 32]> {
        self.take(32)?
            .try_into()
            .map_err(|_| invalid("digest width"))
    }
    fn payload(&mut self) -> Result<&'a [u8]> {
        let length = self.length()?;
        if length > maximum() {
            return Err(limit(
                crate::error::CodecLimitKind::PatchBytes,
                length,
                maximum(),
            ));
        }
        self.take(length)
    }
}

fn read_order(input: &mut Input<'_>, count: usize) -> Result<Vec<StableSlideKey>> {
    let mut order = Vec::new();
    order.try_reserve_exact(count).map_err(allocation)?;
    let mut seen = std::collections::HashSet::new();
    seen.try_reserve(count).map_err(allocation)?;
    for _ in 0..count {
        let key = read_key(input)?;
        if !seen.insert(key.clone()) {
            return Err(invalid("root order repeats a stable key"));
        }
        order.push(key);
    }
    Ok(order)
}

fn read_key(input: &mut Input<'_>) -> Result<StableSlideKey> {
    let uri_len = input.count()?;
    let uri = String::from_utf8(input.take(uri_len)?.to_vec())
        .map_err(|_| invalid("stable key URI is not UTF-8"))?;
    let slide_id = u32::from_le_bytes(
        input
            .take(4)?
            .try_into()
            .map_err(|_| invalid("slide ID width"))?,
    );
    let relationship_len = input.count()?;
    let relationship_id = String::from_utf8(input.take(relationship_len)?.to_vec())
        .map_err(|_| invalid("stable key relationship ID is not UTF-8"))?;
    let part_uri = litchi_opc::PackURI::new(uri).map_err(|_| invalid("invalid stable key URI"))?;
    Ok(stable_slide_key(part_uri, slide_id, relationship_id))
}
