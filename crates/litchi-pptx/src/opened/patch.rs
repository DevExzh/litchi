//! Durable exact-source patches, disjoint joins, and bounded history.

use std::collections::VecDeque;
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};

use super::model::{Limits, Snapshot, capture, invalid, part_context};
use crate::{Error, Result};

const MAGIC: &[u8; 8] = b"LPTX0001";
const CONTEXT_BYTES: usize = 32;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Delta {
    name: PackURI,
    context: [u8; CONTEXT_BYTES],
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
}

impl Delta {
    pub(crate) fn capture(
        source: &OpcPackage,
        target: &OpcPackage,
        name: PackURI,
    ) -> Result<Option<Self>> {
        let before = source.get_part(&name)?;
        let after = target.get_part(&name)?;
        if before.content_type() != after.content_type()
            || part_context(before)? != part_context(after)?
        {
            return Err(invalid(
                "opened-presentation transaction changed unsupported part dependencies",
            ));
        }
        if before.blob() == after.blob() {
            return Ok(None);
        }
        Ok(Some(Self {
            name,
            context: part_context(before)?,
            before: before.blob_arc(),
            after: after.blob_arc(),
        }))
    }

    fn inverse(&self) -> Self {
        Self {
            name: self.name.clone(),
            context: self.context,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// Durable, exact-source patch over a finite set of presentation parts.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    pub(crate) presentation_name: PackURI,
    pub(crate) deltas: Vec<Delta>,
    pub(crate) limits: Limits,
}

impl Patch {
    pub(crate) fn new(
        presentation_name: PackURI,
        mut deltas: Vec<Delta>,
        limits: Limits,
    ) -> Result<Self> {
        deltas.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        if deltas.len() > limits.max_parts() {
            return Err(Error::Limit {
                resource: "opened-presentation patch parts",
                limit: limits.max_parts(),
            });
        }
        if deltas.windows(2).any(|pair| pair[0].name == pair[1].name) {
            return Err(invalid(
                "opened-presentation patch contains overlapping part deltas",
            ));
        }
        let patch = Self {
            presentation_name,
            deltas,
            limits,
        };
        patch.require_encoded_limit()?;
        Ok(patch)
    }

    /// Whether this patch changes no part bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.deltas.is_empty()
    }

    /// Number of exact part resources in the write set.
    #[must_use]
    pub fn resource_count(&self) -> usize {
        self.deltas.len()
    }

    /// Physical part names in deterministic order.
    #[must_use]
    pub fn resources(&self) -> impl ExactSizeIterator<Item = &PackURI> {
        self.deltas.iter().map(|delta| &delta.name)
    }

    /// Exact inverse, suitable only after this patch has been applied.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            presentation_name: self.presentation_name.clone(),
            deltas: self.deltas.iter().map(Delta::inverse).collect(),
            limits: self.limits,
        }
    }

    /// Whether two patches overlap on any physical resource.
    #[must_use]
    pub fn conflicts_with(&self, other: &Self) -> bool {
        self.deltas
            .iter()
            .any(|left| other.deltas.iter().any(|right| left.name == right.name))
    }

    /// Join two disjoint patches into one atomic write set.
    ///
    /// # Errors
    ///
    /// Returns an error for different presentation roots, overlapping parts,
    /// or a combined durable representation beyond the finite limits.
    pub fn join(&self, other: &Self) -> Result<Self> {
        if self.presentation_name != other.presentation_name {
            return Err(invalid(
                "opened-presentation patches belong to different roots",
            ));
        }
        if self.conflicts_with(other) {
            return Err(invalid(
                "opened-presentation patches conflict on a physical part",
            ));
        }
        let limits = intersect_limits(self.limits, other.limits)?;
        let mut deltas = Vec::new();
        deltas
            .try_reserve_exact(self.deltas.len().saturating_add(other.deltas.len()))
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation joined patch",
                source,
            })?;
        deltas.extend(self.deltas.iter().cloned());
        deltas.extend(other.deltas.iter().cloned());
        Self::new(self.presentation_name.clone(), deltas, limits)
    }

    /// Serialize this patch into the stable `LPTX0001` binary format.
    ///
    /// # Errors
    ///
    /// Returns an error if lengths overflow or exceed the configured bound.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let length = self.encoded_len()?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(length)
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation durable patch",
                source,
            })?;
        output.extend_from_slice(MAGIC);
        put_bytes32(&mut output, self.presentation_name.as_str().as_bytes())?;
        put_u32(&mut output, self.deltas.len())?;
        for delta in &self.deltas {
            put_bytes32(&mut output, delta.name.as_str().as_bytes())?;
            output.extend_from_slice(&delta.context);
            put_bytes64(&mut output, delta.before.as_slice())?;
            put_bytes64(&mut output, delta.after.as_slice())?;
        }
        if output.len() != length {
            return Err(invalid(
                "opened-presentation durable patch length changed during encoding",
            ));
        }
        Ok(output)
    }

    /// Parse a stable durable patch under conservative finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid version, malformed length, duplicate
    /// resource, invalid part name, trailing data, or exceeded bound.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a stable durable patch under caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed or unbounded input.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        if bytes.len() > limits.max_patch_bytes() {
            return Err(Error::Limit {
                resource: "opened-presentation durable patch bytes",
                limit: limits.max_patch_bytes(),
            });
        }
        let mut input = Input::new(bytes);
        if input.take(MAGIC.len())? != MAGIC {
            return Err(invalid(
                "opened-presentation durable patch has an unsupported version",
            ));
        }
        let presentation_name = parse_part_name(input.bytes32()?)?;
        let count = input.usize32()?;
        if count > limits.max_parts() {
            return Err(Error::Limit {
                resource: "opened-presentation patch parts",
                limit: limits.max_parts(),
            });
        }
        let mut deltas = Vec::new();
        deltas
            .try_reserve_exact(count)
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation decoded deltas",
                source,
            })?;
        for _index in 0..count {
            let name = parse_part_name(input.bytes32()?)?;
            let context: [u8; CONTEXT_BYTES] = input
                .take(CONTEXT_BYTES)?
                .try_into()
                .map_err(|_err| invalid("opened-presentation patch context is malformed"))?;
            let before = input.bytes64(limits.max_patch_bytes())?.to_vec();
            let after = input.bytes64(limits.max_patch_bytes())?.to_vec();
            if before == after {
                return Err(invalid(
                    "opened-presentation durable patch contains a no-op delta",
                ));
            }
            deltas.push(Delta {
                name,
                context,
                before: Arc::new(before),
                after: Arc::new(after),
            });
        }
        if !input.is_empty() {
            return Err(invalid(
                "opened-presentation durable patch has trailing bytes",
            ));
        }
        Self::new(presentation_name, deltas, limits)
    }

    pub(crate) fn encoded_len(&self) -> Result<usize> {
        let mut length = MAGIC.len();
        length = add(length, 4)?;
        length = add(length, self.presentation_name.as_str().len())?;
        length = add(length, 4)?;
        for delta in &self.deltas {
            length = add(length, 4)?;
            length = add(length, delta.name.as_str().len())?;
            length = add(length, CONTEXT_BYTES)?;
            length = add(length, 8)?;
            length = add(length, delta.before.len())?;
            length = add(length, 8)?;
            length = add(length, delta.after.len())?;
        }
        if length > self.limits.max_patch_bytes() {
            return Err(Error::Limit {
                resource: "opened-presentation durable patch bytes",
                limit: self.limits.max_patch_bytes(),
            });
        }
        Ok(length)
    }

    fn require_encoded_limit(&self) -> Result<()> {
        self.encoded_len().map(|_length| ())
    }
}

/// FIFO undo history bounded by both entry count and encoded bytes.
#[derive(Debug, Clone)]
pub struct History {
    limits: Limits,
    entries: VecDeque<(Patch, usize)>,
    bytes: usize,
}

impl History {
    /// Construct an empty bounded history.
    #[must_use]
    pub fn new(limits: Limits) -> Self {
        Self {
            limits,
            entries: VecDeque::new(),
            bytes: 0,
        }
    }

    /// Number of retained forward patches.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether no undo entry is retained.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Aggregate durable bytes retained by the history.
    #[must_use]
    pub const fn encoded_bytes(&self) -> usize {
        self.bytes
    }

    /// Retain a changed patch, evicting oldest entries to satisfy both bounds.
    ///
    /// # Errors
    ///
    /// Returns an error when one patch alone exceeds the history byte bound.
    pub fn push(&mut self, patch: Patch) -> Result<()> {
        if patch.is_empty() {
            return Ok(());
        }
        let length = patch.encoded_len()?;
        if length > self.limits.max_history_bytes() {
            return Err(Error::Limit {
                resource: "opened-presentation history bytes",
                limit: self.limits.max_history_bytes(),
            });
        }
        while self.entries.len() >= self.limits.max_history_entries()
            || self
                .bytes
                .checked_add(length)
                .is_none_or(|total| total > self.limits.max_history_bytes())
        {
            let Some((_evicted, removed)) = self.entries.pop_front() else {
                break;
            };
            self.bytes = self.bytes.saturating_sub(removed);
        }
        self.bytes = self
            .bytes
            .checked_add(length)
            .ok_or_else(|| invalid("opened-presentation history byte count overflow"))?;
        self.entries.push_back((patch, length));
        Ok(())
    }

    /// Remove the newest entry and return its exact inverse.
    pub fn pop_inverse(&mut self) -> Option<Patch> {
        let (patch, length) = self.entries.pop_back()?;
        self.bytes = self.bytes.saturating_sub(length);
        Some(patch.inverse())
    }
}

pub(crate) fn apply(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let current_main = crate::parts::PresentationPart::from_package(package)?
        .part()
        .partname()
        .clone();
    if current_main != patch.presentation_name {
        return Err(invalid(
            "opened-presentation patch targets a different presentation root",
        ));
    }
    for delta in &patch.deltas {
        let part = package.get_part(&delta.name)?;
        if part_context(part)? != delta.context || part.blob() != delta.before.as_slice() {
            return Err(Error::UnsafeEdit {
                operation: "apply_opened_presentation_patch",
                reason: "an opened-presentation patch resource is stale",
            });
        }
    }
    if patch.is_empty() {
        return capture(package, patch.limits);
    }
    for delta in &patch.deltas {
        package
            .get_part_mut(&delta.name)?
            .set_blob_shared(Arc::clone(&delta.after));
    }
    let snapshot = capture(package, patch.limits)?;
    for delta in &patch.deltas {
        let part = package.get_part(&delta.name)?;
        if part_context(part)? != delta.context || part.blob() != delta.after.as_slice() {
            return Err(invalid(
                "opened-presentation published resource differs from its patch target",
            ));
        }
    }
    package.unsign();
    Ok(snapshot)
}

fn intersect_limits(left: Limits, right: Limits) -> Result<Limits> {
    Limits::new(
        left.max_parts().min(right.max_parts()),
        left.max_patch_bytes().min(right.max_patch_bytes()),
        left.max_text_bytes().min(right.max_text_bytes()),
        left.max_history_entries().min(right.max_history_entries()),
        left.max_history_bytes().min(right.max_history_bytes()),
    )
    .ok_or_else(|| invalid("opened-presentation patch limits are invalid"))
}

fn put_u32(output: &mut Vec<u8>, value: usize) -> Result<()> {
    output.extend_from_slice(
        &u32::try_from(value)
            .map_err(|_err| invalid("opened-presentation durable length exceeds u32"))?
            .to_le_bytes(),
    );
    Ok(())
}

fn put_bytes32(output: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    put_u32(output, value.len())?;
    output.extend_from_slice(value);
    Ok(())
}

fn put_bytes64(output: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    output.extend_from_slice(
        &u64::try_from(value.len())
            .map_err(|_err| invalid("opened-presentation durable length exceeds u64"))?
            .to_le_bytes(),
    );
    output.extend_from_slice(value);
    Ok(())
}

fn add(left: usize, right: usize) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| invalid("opened-presentation durable patch length overflow"))
}

fn parse_part_name(value: &[u8]) -> Result<PackURI> {
    let text = std::str::from_utf8(value)
        .map_err(|_err| invalid("opened-presentation patch part name is not UTF-8"))?;
    PackURI::new(text).map_err(Error::Invalid)
}

struct Input<'a> {
    bytes: &'a [u8],
    position: usize,
}

impl<'a> Input<'a> {
    const fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, position: 0 }
    }

    fn take(&mut self, length: usize) -> Result<&'a [u8]> {
        let end = self
            .position
            .checked_add(length)
            .ok_or_else(|| invalid("opened-presentation patch offset overflow"))?;
        let value = self
            .bytes
            .get(self.position..end)
            .ok_or_else(|| invalid("opened-presentation durable patch is truncated"))?;
        self.position = end;
        Ok(value)
    }

    fn usize32(&mut self) -> Result<usize> {
        let raw: [u8; 4] = self
            .take(4)?
            .try_into()
            .map_err(|_err| invalid("opened-presentation u32 is malformed"))?;
        usize::try_from(u32::from_le_bytes(raw))
            .map_err(|_err| invalid("opened-presentation u32 exceeds usize"))
    }

    fn usize64(&mut self) -> Result<usize> {
        let raw: [u8; 8] = self
            .take(8)?
            .try_into()
            .map_err(|_err| invalid("opened-presentation u64 is malformed"))?;
        usize::try_from(u64::from_le_bytes(raw))
            .map_err(|_err| invalid("opened-presentation u64 exceeds usize"))
    }

    fn bytes32(&mut self) -> Result<&'a [u8]> {
        let length = self.usize32()?;
        self.take(length)
    }

    fn bytes64(&mut self, limit: usize) -> Result<&'a [u8]> {
        let length = self.usize64()?;
        if length > limit {
            return Err(Error::Limit {
                resource: "opened-presentation durable resource bytes",
                limit,
            });
        }
        self.take(length)
    }

    fn is_empty(&self) -> bool {
        self.position == self.bytes.len()
    }
}
