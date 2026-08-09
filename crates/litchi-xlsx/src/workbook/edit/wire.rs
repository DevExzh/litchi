//! Bounded deterministic durable wire for ordinary workbook patches.

use std::collections::BTreeMap;

use litchi_core::patch::{
    BlobBundle, BlobId, BlobLimits, ForwardOnly, Patch as CorePatch, PatchLimits, PatchOperation,
    Reversible, ReversibleOperation,
};
use serde_json::Value;

use super::model::Patch;
use crate::error::{Error, Result, allocation, invalid};
use crate::workbook::Workbook;

const MAX_PACKAGE_BYTES: usize = 64 * 1024 * 1024;
const MAX_WIRE_JSON_BYTES: usize = 192 * 1024 * 1024;
const FORMAT: &str = "litchi.xlsx.workbook";
const OPERATION: &str = "workbook.replace";
const TARGET: &str = "/workbook";
const SOURCE_PRECONDITION: &str = "source_sha256";

/// Reversible XLSX workbook patch encoded through the common durable wire.
#[derive(Clone)]
pub struct DurablePatch {
    inner: CorePatch<Reversible>,
}

impl DurablePatch {
    pub(super) fn from_patch(patch: &Patch) -> Result<Self> {
        let before = patch
            .source
            .as_ref()
            .ok_or_else(|| invalid("workbook patch has no durable source snapshot"))?;
        let after = patch
            .target
            .as_ref()
            .ok_or_else(|| invalid("workbook patch has no durable target snapshot"))?;
        Self::from_workbooks(before, after)
    }

    fn from_workbooks(before: &Workbook, after: &Workbook) -> Result<Self> {
        let before = workbook_bytes(before)?;
        let after = workbook_bytes(after)?;
        validate_workbook_bytes(&before)?;
        validate_workbook_bytes(&after)?;

        let limits = patch_limits();
        let mut forward_blobs = BlobBundle::new(limits.blobs());
        let after_id = forward_blobs.insert(&after)?;
        let mut reverse_blobs = BlobBundle::new(limits.blobs());
        let before_id = reverse_blobs.insert(&before)?;
        let forward = operation(limits, &before_id, &after_id)?;
        let inverse = operation(limits, &after_id, &before_id)?;
        let inner = CorePatch::<Reversible>::new(
            limits,
            FORMAT,
            [ReversibleOperation::new(forward, inverse)],
            forward_blobs,
            reverse_blobs,
        )?;
        Ok(Self { inner })
    }

    /// Parse canonical deterministic JSON under XLSX's finite wire bounds.
    ///
    /// # Errors
    ///
    /// Returns an error for non-canonical JSON, invalid blob integrity,
    /// exceeded limits, a foreign vocabulary, or invalid workbook artifacts.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<Reversible>::from_deterministic_json(bytes, patch_limits())?;
        validate_reversible(&inner)?;
        Ok(Self { inner })
    }

    /// Serialize canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error if the bounded wire output cannot be produced.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        Ok(self.inner.to_deterministic_json()?)
    }

    /// Apply only when `source` is byte-identical to the expected workbook.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source bytes or an invalid target workbook.
    pub fn apply(&self, source: &Workbook) -> Result<Workbook> {
        let inverse = self.inner.inverse();
        let expected = direction(&inverse)?.target_bytes;
        if workbook_bytes(source)? != expected {
            return Err(stale());
        }
        workbook_from_target(&self.inner)
    }

    /// Return the exact durable inverse without copying attached blobs.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: self.inner.inverse(),
        }
    }

    /// Permanently discard the inverse operation and source workbook blob.
    #[must_use]
    pub fn seal(self) -> SealedPatch {
        SealedPatch {
            inner: self.inner.seal(),
        }
    }
}

/// Forward-only durable XLSX workbook patch.
#[derive(Clone)]
pub struct SealedPatch {
    inner: CorePatch<ForwardOnly>,
}

impl SealedPatch {
    /// Parse canonical forward-only deterministic JSON under finite bounds.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid envelope, vocabulary, or target XLSX.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<ForwardOnly>::from_deterministic_json(bytes, patch_limits())?;
        validate_sealed(&inner)?;
        Ok(Self { inner })
    }

    /// Serialize canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error if the bounded wire output cannot be produced.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        Ok(self.inner.to_deterministic_json()?)
    }

    /// Apply after checking the retained SHA-256 source precondition.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale source digest or invalid target workbook.
    pub fn apply(&self, source: &Workbook) -> Result<Workbook> {
        let direction = direction(&self.inner)?;
        let source_bytes = workbook_bytes(source)?;
        if BlobId::of(&source_bytes).as_hex() != direction.source_id {
            return Err(stale());
        }
        workbook_from_target(&self.inner)
    }
}

struct Direction<'a> {
    source_id: &'a str,
    target_id: &'a str,
    target_bytes: &'a [u8],
}

fn patch_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(1, MAX_PACKAGE_BYTES, MAX_PACKAGE_BYTES),
        MAX_WIRE_JSON_BYTES,
        1,
        4,
        4_096,
        16_384,
    )
}

fn operation(limits: PatchLimits, source: &BlobId, target: &BlobId) -> Result<PatchOperation> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        SOURCE_PRECONDITION.to_owned(),
        Value::String(source.as_hex()),
    );
    Ok(PatchOperation::new(
        limits,
        OPERATION,
        TARGET,
        preconditions,
        Value::String(target.as_hex()),
    )?)
}

fn direction<Mode>(patch: &CorePatch<Mode>) -> Result<Direction<'_>> {
    if patch.format() != FORMAT || patch.operations().len() != 1 {
        return Err(invalid_vocabulary());
    }
    let operation = &patch.operations()[0];
    if operation.op != OPERATION || operation.target != TARGET || operation.preconditions.len() != 1
    {
        return Err(invalid_vocabulary());
    }
    let source_id = operation
        .preconditions
        .get(SOURCE_PRECONDITION)
        .and_then(Value::as_str)
        .ok_or_else(invalid_vocabulary)?;
    let target_id = operation.value.as_str().ok_or_else(invalid_vocabulary)?;
    if !canonical_digest(source_id) || !canonical_digest(target_id) || patch.blobs().len() != 1 {
        return Err(invalid_vocabulary());
    }
    let id = patch.blobs().ids().next().ok_or_else(invalid_vocabulary)?;
    if id.as_hex() != target_id {
        return Err(invalid_vocabulary());
    }
    let target_bytes = patch.blobs().get(id).ok_or_else(invalid_vocabulary)?;
    Ok(Direction {
        source_id,
        target_id,
        target_bytes,
    })
}

fn validate_reversible(patch: &CorePatch<Reversible>) -> Result<()> {
    let forward = direction(patch)?;
    let inverse = patch.inverse();
    let reverse = direction(&inverse)?;
    if forward.source_id != reverse.target_id || forward.target_id != reverse.source_id {
        return Err(invalid_vocabulary());
    }
    validate_workbook_bytes(forward.target_bytes)?;
    validate_workbook_bytes(reverse.target_bytes)
}

fn validate_sealed(patch: &CorePatch<ForwardOnly>) -> Result<()> {
    validate_workbook_bytes(direction(patch)?.target_bytes)
}

fn workbook_from_target<Mode>(patch: &CorePatch<Mode>) -> Result<Workbook> {
    let target = direction(patch)?.target_bytes;
    Workbook::from_bytes(copy_bytes(target)?)
}

fn validate_workbook_bytes(bytes: &[u8]) -> Result<()> {
    let _ = Workbook::from_bytes(copy_bytes(bytes)?)?;
    Ok(())
}

fn workbook_bytes(workbook: &Workbook) -> Result<Vec<u8>> {
    let bytes = workbook.to_plain_bytes()?;
    ensure_size(bytes.len())?;
    Ok(bytes)
}

fn copy_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    ensure_size(bytes.len())?;
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|source| allocation("durable workbook bytes", source))?;
    copy.extend_from_slice(bytes);
    Ok(copy)
}

fn ensure_size(size: usize) -> Result<()> {
    if size > MAX_PACKAGE_BYTES {
        return Err(invalid(format!(
            "durable workbook exceeds the {MAX_PACKAGE_BYTES}-byte package limit"
        )));
    }
    Ok(())
}

fn canonical_digest(value: &str) -> bool {
    value.len() == 64
        && value
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
}

fn invalid_vocabulary() -> Error {
    invalid("invalid durable XLSX workbook patch vocabulary")
}

fn stale() -> Error {
    Error::PatchConflict {
        part: "/".to_owned(),
    }
}
