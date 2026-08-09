//! Durable semantic ODB patch envelopes.

use std::collections::BTreeMap;

use litchi_core::{
    BlobBundle, BlobId, BlobLimits, Error, ForwardOnly, Patch as CorePatch, PatchLimits,
    PatchOperation, Result, Reversible, ReversibleOperation,
};
use serde_json::Value;

use super::{Change, ChangeAction, ChangeKind, Patch};
use crate::Database;

const FORMAT: &str = "litchi.odb";
const SOURCE_PRECONDITION: &str = "source_sha256";
const MAX_DURABLE_PACKAGE_BYTES: usize = 32 * 1024 * 1024;
const MAX_WIRE_JSON_BYTES: usize = 96 * 1024 * 1024;
const MAX_DURABLE_OPERATIONS: usize = 65_536;
const MAX_OPERATION_PAYLOAD_BYTES: usize = 16 * 1024 * 1024;

/// Bounded deterministic-JSON reversible ODB patch.
#[derive(Clone)]
pub struct DurablePatch {
    inner: CorePatch<Reversible>,
}

impl DurablePatch {
    fn from_patch(patch: &Patch) -> Result<Self> {
        ensure_package_size(patch.source.as_bytes())?;
        ensure_package_size(patch.target.as_bytes())?;
        let limits = durable_limits();
        let mut forward_blobs = BlobBundle::new(limits.blobs());
        let target_id = forward_blobs
            .insert(patch.target.as_bytes())
            .map_err(wire_error)?;
        let mut reverse_blobs = BlobBundle::new(limits.blobs());
        let source_id = reverse_blobs
            .insert(patch.source.as_bytes())
            .map_err(wire_error)?;
        let operations = reversible_operations(
            &patch.changes,
            &source_id.as_hex(),
            &target_id.as_hex(),
            limits,
        )?;
        let inner =
            CorePatch::<Reversible>::new(limits, FORMAT, operations, forward_blobs, reverse_blobs)
                .map_err(wire_error)?;
        Ok(Self { inner })
    }

    /// Parses canonical deterministic JSON under ODB's finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error for a noncanonical envelope, a foreign vocabulary,
    /// invalid blob integrity, exceeded bounds, or invalid ODB artifacts.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<Reversible>::from_deterministic_json(bytes, durable_limits())
            .map_err(wire_error)?;
        validate_reversible(&inner)?;
        Ok(Self { inner })
    }

    /// Serializes this patch as canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error if bounded serialization fails.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner.to_deterministic_json().map_err(wire_error)
    }

    /// Applies the patch only to its exact retained source artifact.
    ///
    /// # Errors
    ///
    /// Returns an error on stale source bytes or invalid target readback.
    pub fn apply(&self, source: &Database) -> Result<Database> {
        let reverse = self.inner.inverse();
        let expected = direction(&reverse)?.artifact;
        if source.as_bytes() != expected {
            return source_mismatch();
        }
        database_from_target(&self.inner)
    }

    /// Returns a durable patch that restores the exact source artifact.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: self.inner.inverse(),
        }
    }

    /// Permanently discards reverse operations and source package bytes.
    #[must_use]
    pub fn seal(self) -> SealedPatch {
        SealedPatch {
            inner: self.inner.seal(),
        }
    }
}

/// Forward-only durable ODB patch.
#[derive(Clone)]
pub struct SealedPatch {
    inner: CorePatch<ForwardOnly>,
}

impl SealedPatch {
    /// Parses a canonical bounded forward-only ODB patch.
    ///
    /// # Errors
    ///
    /// Returns an error for a foreign vocabulary, invalid artifact, integrity
    /// failure, noncanonical JSON, or exceeded limit.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<ForwardOnly>::from_deterministic_json(bytes, durable_limits())
            .map_err(wire_error)?;
        validate_direction(&inner)?;
        Ok(Self { inner })
    }

    /// Serializes this patch as canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error if bounded serialization fails.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner.to_deterministic_json().map_err(wire_error)
    }

    /// Applies after checking the retained cryptographic source precondition.
    ///
    /// # Errors
    ///
    /// Returns an error on a stale source digest or invalid target readback.
    pub fn apply(&self, source: &Database) -> Result<Database> {
        let selected = direction(&self.inner)?;
        if BlobId::of(source.as_bytes()).as_hex() != selected.source_id {
            return source_mismatch();
        }
        reopen_artifact(selected.artifact)
    }
}

impl Patch {
    /// Converts this in-memory exact patch into a deterministic semantic wire
    /// envelope with content-addressed package artifacts.
    ///
    /// # Errors
    ///
    /// Returns an error when an artifact or operation exceeds durable bounds.
    pub fn durable(&self) -> Result<DurablePatch> {
        DurablePatch::from_patch(self)
    }
}

struct Direction<'a> {
    source_id: String,
    target_id: String,
    artifact: &'a [u8],
}

fn durable_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(1, MAX_DURABLE_PACKAGE_BYTES, MAX_DURABLE_PACKAGE_BYTES),
        MAX_WIRE_JSON_BYTES,
        MAX_DURABLE_OPERATIONS,
        5,
        4_096,
        MAX_OPERATION_PAYLOAD_BYTES,
    )
}

fn reversible_operations(
    changes: &[Change],
    source_id: &str,
    target_id: &str,
    limits: PatchLimits,
) -> Result<Vec<ReversibleOperation>> {
    let count = changes.len().max(1);
    let mut operations = Vec::new();
    operations
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "ODB durable operations",
            source,
        })?;
    if changes.is_empty() {
        let forward = operation("odb.noop", "database", source_id, target_id, limits)?;
        let inverse = operation("odb.noop", "database", target_id, source_id, limits)?;
        operations.push(ReversibleOperation::new(forward, inverse));
        return Ok(operations);
    }
    for change in changes {
        let forward = operation(
            &operation_name(change.kind(), change.action()),
            change.target(),
            source_id,
            target_id,
            limits,
        )?;
        let inverse = operation(
            &operation_name(change.kind(), inverse_action(change.action())),
            change.target(),
            target_id,
            source_id,
            limits,
        )?;
        operations.push(ReversibleOperation::new(forward, inverse));
    }
    Ok(operations)
}

fn operation(
    name: &str,
    target: &str,
    source_id: &str,
    target_id: &str,
    limits: PatchLimits,
) -> Result<PatchOperation> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        SOURCE_PRECONDITION.to_string(),
        Value::String(source_id.to_owned()),
    );
    PatchOperation::new(
        limits,
        name,
        target,
        preconditions,
        Value::String(target_id.to_owned()),
    )
    .map_err(wire_error)
}

fn direction<Mode>(patch: &CorePatch<Mode>) -> Result<Direction<'_>> {
    if patch.format() != FORMAT || patch.operations().is_empty() || patch.blobs().len() != 1 {
        return invalid_wire();
    }
    let mut source_id = None;
    let mut target_id = None;
    for operation in patch.operations() {
        if !valid_operation_name(&operation.op)
            || operation.preconditions.len() != 1
            || operation.target.is_empty()
        {
            return invalid_wire();
        }
        let source = operation
            .preconditions
            .get(SOURCE_PRECONDITION)
            .and_then(Value::as_str)
            .ok_or_else(invalid_wire_error)?;
        let target = operation.value.as_str().ok_or_else(invalid_wire_error)?;
        if !is_digest(source) || !is_digest(target) {
            return invalid_wire();
        }
        if source_id.as_deref().is_some_and(|value| value != source)
            || target_id.as_deref().is_some_and(|value| value != target)
        {
            return invalid_wire();
        }
        source_id = Some(source.to_owned());
        target_id = Some(target.to_owned());
    }
    let source_id = source_id.ok_or_else(invalid_wire_error)?;
    let target_id = target_id.ok_or_else(invalid_wire_error)?;
    let blob_id = patch.blobs().ids().next().ok_or_else(invalid_wire_error)?;
    if blob_id.as_hex() != target_id {
        return invalid_wire();
    }
    let artifact = patch.blobs().get(blob_id).ok_or_else(invalid_wire_error)?;
    Ok(Direction {
        source_id,
        target_id,
        artifact,
    })
}

fn validate_reversible(patch: &CorePatch<Reversible>) -> Result<()> {
    let forward = direction(patch)?;
    let reverse_patch = patch.inverse();
    let reverse = direction(&reverse_patch)?;
    if forward.source_id != reverse.target_id || forward.target_id != reverse.source_id {
        return invalid_wire();
    }
    reopen_artifact(forward.artifact)?;
    reopen_artifact(reverse.artifact)?;
    Ok(())
}

fn validate_direction<Mode>(patch: &CorePatch<Mode>) -> Result<()> {
    let selected = direction(patch)?;
    reopen_artifact(selected.artifact)?;
    Ok(())
}

fn database_from_target<Mode>(patch: &CorePatch<Mode>) -> Result<Database> {
    let selected = direction(patch)?;
    reopen_artifact(selected.artifact)
}

fn reopen_artifact(bytes: &[u8]) -> Result<Database> {
    let database = Database::from_bytes(copy_bytes(bytes)?)?;
    database.catalog()?;
    Ok(database)
}

fn operation_name(kind: ChangeKind, action: ChangeAction) -> String {
    let kind = match kind {
        ChangeKind::Connection => "connection",
        ChangeKind::Query => "query",
        ChangeKind::Table => "table",
        ChangeKind::Column => "column",
        ChangeKind::Key => "key",
        ChangeKind::Index => "index",
        ChangeKind::Component => "component",
        ChangeKind::ProducerExtension => "producer-extension",
    };
    let action = match action {
        ChangeAction::Create => "create",
        ChangeAction::Update => "update",
        ChangeAction::Remove => "remove",
    };
    format!("odb.{kind}.{action}")
}

fn inverse_action(action: ChangeAction) -> ChangeAction {
    match action {
        ChangeAction::Create => ChangeAction::Remove,
        ChangeAction::Update => ChangeAction::Update,
        ChangeAction::Remove => ChangeAction::Create,
    }
}

fn valid_operation_name(value: &str) -> bool {
    if value == "odb.noop" {
        return true;
    }
    let mut parts = value.split('.');
    matches!(parts.next(), Some("odb"))
        && matches!(
            parts.next(),
            Some(
                "connection"
                    | "query"
                    | "table"
                    | "column"
                    | "key"
                    | "index"
                    | "component"
                    | "producer-extension"
            )
        )
        && matches!(parts.next(), Some("create" | "update" | "remove"))
        && parts.next().is_none()
}

fn is_digest(value: &str) -> bool {
    value.len() == 64
        && value
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
}

fn ensure_package_size(bytes: &[u8]) -> Result<()> {
    if bytes.len() > MAX_DURABLE_PACKAGE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODB durable package exceeds the {MAX_DURABLE_PACKAGE_BYTES}-byte limit"
        )));
    }
    Ok(())
}

fn copy_bytes(source: &[u8]) -> Result<Vec<u8>> {
    ensure_package_size(source)?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "ODB durable package",
            source,
        })?;
    bytes.extend_from_slice(source);
    Ok(bytes)
}

fn wire_error(source: litchi_core::PatchError) -> Error {
    let message = format!("invalid ODB durable patch: {source}");
    drop(source);
    Error::InvalidFormat(message)
}

fn invalid_wire<T>() -> Result<T> {
    Err(invalid_wire_error())
}

fn invalid_wire_error() -> Error {
    Error::InvalidFormat("invalid ODB durable patch vocabulary".to_string())
}

fn source_mismatch<T>() -> Result<T> {
    Err(Error::InvalidFormat(
        "ODB durable patch source does not match".to_string(),
    ))
}
