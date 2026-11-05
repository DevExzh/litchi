//! Persisted-record rewrites staged by element edits (record 0323).
//!
//! Comments, hyperlinks and table-cell text live in records the slide-order
//! root does not otherwise model: the slide's `___PPT10` programmable tag, a
//! shape's `OfficeArtClientData` and `ClientTextbox`, and the live document's
//! `ExObjListContainer`. An element edit computes the complete replacement
//! of each record it changes from the exact source bytes of that record; this
//! module publishes those replacements in one append-only user edit, proves
//! that every replaced record reads back byte-exactly and that the slide list
//! is unchanged, and carries the change through exact, inverse and durable
//! patches. Every record an edit does not name keeps its bytes.

use std::sync::Arc;

use super::{
    Error, PackageError, Position, Result, Snapshot, persisted_object_record, persisted_record,
};

/// The persisted owner one record rewrite replaces.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Owner {
    /// The live `DocumentContainer`.
    Document,
    /// The slide at this position of the artifact's own slide list.
    Slide(Position),
    /// The persisted record with this persist identifier, such as a master.
    Persist(u32),
}

/// One complete persisted-record replacement.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Rewrite {
    pub(crate) owner: Owner,
    pub(crate) before: Arc<[u8]>,
    pub(crate) after: Arc<[u8]>,
}

/// The record rewrites of one element edit, published together.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ElementChange {
    pub(crate) rewrites: Vec<Rewrite>,
}

impl ElementChange {
    /// The change that restores every record, in reverse order.
    pub(crate) fn inverse(&self) -> Self {
        Self {
            rewrites: self
                .rewrites
                .iter()
                .rev()
                .map(|rewrite| Rewrite {
                    owner: rewrite.owner,
                    before: Arc::clone(&rewrite.after),
                    after: Arc::clone(&rewrite.before),
                })
                .collect(),
        }
    }
}

/// The records an element edit reads: the exact persisted bytes of the
/// selected slide and, when requested, the live document.
pub(crate) struct View<'a> {
    /// The selected slide's `SlideContainer`.
    pub(crate) slide: Option<SlideRecord<'a>>,
    /// The live `DocumentContainer`.
    pub(crate) document: Option<&'a [u8]>,
}

/// One slide record offered to an element edit.
pub(crate) struct SlideRecord<'a> {
    pub(crate) bytes: &'a [u8],
}

/// The replacements an element edit asks for; `None` keeps the record.
#[derive(Debug, Default)]
pub(crate) struct Replacement {
    pub(crate) slide: Option<Vec<u8>>,
    pub(crate) document: Option<Vec<u8>>,
}

impl Replacement {
    pub(crate) const fn is_noop(&self) -> bool {
        self.slide.is_none() && self.document.is_none()
    }
}

fn owner_persist_id(snapshot: &Snapshot, owner: Owner) -> Result<u32> {
    match owner {
        Owner::Document => Ok(snapshot.document_persist_id),
        Owner::Slide(position) => snapshot
            .document
            .slides()
            .get(position.get())
            .map(|slide| slide.persist_id())
            .ok_or(Error::Refused(super::Refusal::SlideNotFound { position })),
        Owner::Persist(persist_id) => Ok(persist_id),
    }
}

/// The current exact bytes of one owner in `snapshot`.
pub(crate) fn owner_record(snapshot: &Snapshot, owner: Owner) -> Result<Vec<u8>> {
    match owner {
        Owner::Document => Ok(snapshot.document.bytes().to_vec()),
        Owner::Slide(_) => persisted_record(snapshot, owner_persist_id(snapshot, owner)?),
        Owner::Persist(persist_id) => persisted_object_record(snapshot, persist_id),
    }
}

/// Publish `rewrites` over `snapshot` in one append-only user edit.
///
/// Every rewrite's `before` must be the owner's current exact record. The
/// candidate is reopened, its slide list must equal the source's, and every
/// rewritten owner must read back as exactly its `after` bytes.
pub(crate) fn publish(snapshot: &Snapshot, rewrites: &[Rewrite]) -> Result<Snapshot> {
    let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        snapshot.shared_bytes(),
        snapshot.limits.max_package_bytes,
        snapshot.limits,
    )?;
    for rewrite in rewrites {
        let persist_id = owner_persist_id(snapshot, rewrite.owner)?;
        if editor.persisted_record(persist_id)?.as_slice() != &*rewrite.before {
            return Err(PackageError::Corrupted(
                "PPT element record changed before publication".into(),
            )
            .into());
        }
        editor.replace_persisted_record(persist_id, rewrite.after.to_vec())?;
    }
    let bytes = editor.finish()?;
    crate::font::validate_unrelated_streams(snapshot.bytes(), &bytes)?;
    let published = Snapshot::from_bytes_with_options(bytes, snapshot.open_options())?;
    if published.document.slides() != snapshot.document.slides() {
        return Err(PackageError::Corrupted(
            "PPT element publication changed the slide list".into(),
        )
        .into());
    }
    for rewrite in rewrites {
        if owner_record(&published, rewrite.owner)?.as_slice() != &*rewrite.after {
            return Err(PackageError::Corrupted(
                "PPT element record did not read back exactly".into(),
            )
            .into());
        }
    }
    Ok(published)
}

/// The durable operation of one element change: each owner's record SHA-256
/// before the change and its replacement as a blob.
pub(crate) fn durable_operation(
    limits: litchi_core::patch::PatchLimits,
    change: &ElementChange,
    blobs: &mut litchi_core::patch::BlobBundle,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut records = Vec::with_capacity(change.rewrites.len());
    for rewrite in &change.rewrites {
        let blob = blobs.insert_shared(Arc::clone(&rewrite.after))?;
        records.push(serde_json::json!({
            "owner": owner_label(rewrite.owner),
            "before_sha256": super::artifact_hash(&rewrite.before),
            "after": blob.as_hex(),
        }));
    }
    litchi_core::patch::PatchOperation::new(
        limits,
        "presentation.element-records.set",
        "records",
        std::collections::BTreeMap::new(),
        serde_json::Value::Array(records),
    )
}

fn owner_label(owner: Owner) -> String {
    match owner {
        Owner::Document => "document".to_owned(),
        Owner::Slide(position) => format!("slide:{}", position.get()),
        Owner::Persist(persist_id) => format!("persist:{persist_id}"),
    }
}

fn parse_owner(label: &str) -> Result<Owner> {
    if label == "document" {
        return Ok(Owner::Document);
    }
    if let Some(persist_id) = label.strip_prefix("persist:") {
        return persist_id
            .parse::<u32>()
            .ok()
            .map(Owner::Persist)
            .ok_or_else(|| super::invalid_durable_patch("element record owner is malformed"));
    }
    label
        .strip_prefix("slide:")
        .and_then(|position| position.parse::<usize>().ok())
        .map(|position| Owner::Slide(Position::new(position)))
        .ok_or_else(|| super::invalid_durable_patch("element record owner is malformed"))
}

/// Apply one durable element-record operation.
pub(crate) fn apply_durable(
    snapshot: &Snapshot,
    operation: &litchi_core::patch::PatchOperation,
    blobs: &litchi_core::patch::BlobBundle,
) -> Result<Snapshot> {
    if !operation.preconditions.is_empty() || operation.target != "records" {
        return Err(super::invalid_durable_patch(
            "element-record operation has an unexpected target or precondition",
        ));
    }
    let records = operation
        .value
        .as_array()
        .ok_or_else(|| super::invalid_durable_patch("element records must be an array"))?;
    let mut rewrites = Vec::with_capacity(records.len());
    for record in records {
        let field = |name: &str| {
            record
                .get(name)
                .and_then(serde_json::Value::as_str)
                .ok_or_else(|| super::invalid_durable_patch("element record field is missing"))
        };
        let owner = parse_owner(field("owner")?)?;
        let before = owner_record(snapshot, owner)?;
        if super::artifact_hash(&before) != field("before_sha256")? {
            return Err(PackageError::InvalidFormat(
                "PPT durable element-record precondition does not match".into(),
            )
            .into());
        }
        let after = super::blob_by_hex(blobs, field("after")?)
            .ok_or_else(|| super::invalid_durable_patch("element record blob is missing"))?;
        rewrites.push(Rewrite {
            owner,
            before: Arc::from(before),
            after: Arc::from(after),
        });
    }
    if snapshot.has_signature {
        return Err(Error::Refused(super::Refusal::Signed));
    }
    publish(snapshot, &rewrites)
}
