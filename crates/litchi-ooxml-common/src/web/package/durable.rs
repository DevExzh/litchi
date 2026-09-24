//! Durable semantic Web Extensions edits.
//!
//! The wire format in this module carries the typed request that produced an
//! ordinary [`super::Patch`].  Applying an edit decodes that request and sends
//! it through the existing graph planner.  The reverse blob contains only a
//! bounded source closure; it is used to construct a private OPC proof source
//! for inverse replay and is never installed as arbitrary package bytes.
//!
//! The public flow is `Patch::to_durable(PatchLimits)`, deterministic JSON
//! serialization through the common core patch, bounded JSON decoding, and
//! `apply_durable_with_limits` (or the DOCX facade's corresponding method).
//! The source guard covers the retained logical OPC graph, content types,
//! relationship tokens, and affected XML provenance; compressed ZIP artifact
//! identity and unrelated non-part archive entries are outside that logical
//! guard and remain owned by the package source.

use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::fmt::Write as _;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::patch::{
    BlobBundle, BlobId, BlobLimitKind, Patch as CorePatch, PatchError, PatchLimits, PatchOperation,
    Reversible, ReversibleOperation,
};
use litchi_opc::phys_pkg::PhysPkgWriter;
use litchi_opc::{BlobPart, OpcPackage, OwnedRelationships, PackURI};
use serde_json::Value;

use super::super::codec::{
    invalid, parse_add_in_with_budget, parse_panes_with_budget, parse_xml_owned, require_name,
    write_add_in_with, write_panes_with,
};
use super::super::model::{
    AddIn, BackgroundAppData, Conformance, ContainsCustomFunctions, CustomFunctionList,
    CustomFunctions, Dock, ExtList, Limits, OperationBudget, Pane, Panes, Selector,
    SnapshotResource, SnapshotTarget,
};
use super::super::{Error, Result, TASK_PANES_NAMESPACE};
use super::{PartChange, PartState, Patch};

#[path = "durable_scope.rs"]
mod scope;
use scope::scope_hash_from_patch_scope;

#[cfg(test)]
#[path = "durable_intent_tests.rs"]
mod durable_intent_tests;
#[cfg(test)]
#[path = "durable_tests.rs"]
mod durable_tests;

const FORMAT_NAME: &str = "litchi-ooxml-common/web-custom-functions";
const EDIT: &str = "web.edit";
const RESTORE: &str = "web.restore";
const NOOP: &str = "web.noop";
const INTENT_HEADER: &[u8] = b"WCF1";
const RESTORE_HEADER: &[u8] = b"WCR1";
const CLOSURE_HEADER: &[u8] = b"WCC1";
const MAX_INTENT_OPERATIONS: usize = 65_536;
const MAX_INTENT_BYTES: usize = 512 * 1024 * 1024;
const MAX_CLOSURE_RECORDS: usize = 65_536;
const MAX_CLOSURE_BYTES: usize = 512 * 1024 * 1024;
const MAX_MEMBER_NAME: usize = 4096;
const MAX_CONTENT_TYPE: usize = 4096;
const MAX_RELATIONSHIP_BYTES: usize = 8 * 1024 * 1024;
const PROOF_ENTRY_OVERHEAD: usize = 128;

/// How a custom-function edit resolves its graph owners.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GraphMode {
    /// Resolve owners in the existing graph only.
    ExistingOnly,
    /// Retain the existing graph and append authored panes.
    EnsureGraph,
    /// Replace a complete graph. Use [`WebIntent::Put`] to carry that graph.
    ReplaceManagedGraph,
}

/// Semantic custom-function owner selector. Physical relationship IDs remain
/// source-private and are bound by the ordinary patch closure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OwnerSelector {
    AddInId(String),
    PaneIndex(usize),
}

/// One typed custom-function operation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CustomFunctionOperation {
    Set(Option<CustomFunctions>),
    SetContains(Option<ContainsCustomFunctions>),
    SetBackground(Option<BackgroundAppData>),
    SetList(Option<CustomFunctionList>),
    InsertId { index: usize, id: String },
    ReplaceId { index: usize, id: String },
    RemoveId { index: usize },
}

/// A bounded semantic custom-function edit. `additions` is only admitted for
/// [`GraphMode::EnsureGraph`]. Complete graph replacement is represented by a
/// [`WebIntent::Put`] so no panes or resources can be silently omitted.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomFunctionEdit {
    pub mode: GraphMode,
    pub operations: Vec<(OwnerSelector, CustomFunctionOperation)>,
    pub additions: Vec<Pane>,
}

impl CustomFunctionEdit {
    #[must_use]
    pub fn existing(operations: Vec<(OwnerSelector, CustomFunctionOperation)>) -> Self {
        Self {
            mode: GraphMode::ExistingOnly,
            operations,
            additions: Vec::new(),
        }
    }

    #[must_use]
    pub fn ensure(
        additions: Vec<Pane>,
        operations: Vec<(OwnerSelector, CustomFunctionOperation)>,
    ) -> Self {
        Self {
            mode: GraphMode::EnsureGraph,
            operations,
            additions,
        }
    }
}

/// Typed intent retained by an ordinary Web Extensions [`Patch`].
#[derive(Debug, Clone, PartialEq)]
pub enum WebIntent {
    Put {
        panes: Panes,
        conformance: Conformance,
    },
    Remove,
    CustomFunctions(CustomFunctionEdit),
}

/// Attach semantic intent to a planned package transaction.
#[allow(
    dead_code,
    reason = "planner owners bind intent in their integration hook"
)]
pub(in crate::web::package) fn bind_intent(patch: Patch, intent: WebIntent) -> Patch {
    patch.with_durable_intent(intent)
}

/// Build the narrow owner-scoped intent used when the planner proves that a
/// graph change touches only typed custom-function metadata. The complete
/// graph remains source-bound by the ordinary patch, while inverse replay
/// authorizes only these selected add-in owners.
pub(in crate::web::package) fn custom_edit_for_graph(
    existing: &Panes,
    incoming: &Panes,
) -> Result<CustomFunctionEdit> {
    if existing.len() != incoming.len() {
        return invalid("custom-function graph intent changed pane count".into());
    }
    let mut operations = Vec::new();
    operations
        .try_reserve_exact(existing.len())
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions custom-function intents",
            source,
        })?;
    for (old, new) in existing.iter().zip(incoming.iter()) {
        if old.add_in.id != new.add_in.id {
            return invalid("custom-function graph intent changed an owner ID".into());
        }
        if old.add_in.custom_functions() != new.add_in.custom_functions() {
            operations.push((
                OwnerSelector::AddInId(new.add_in.id.clone()),
                CustomFunctionOperation::Set(new.add_in.custom_functions().cloned()),
            ));
        }
    }
    if operations.is_empty() {
        return invalid("custom-function graph intent has no changed owner".into());
    }
    Ok(CustomFunctionEdit::existing(operations))
}

impl Patch {
    /// Bind the typed request that created this source-checked transaction.
    ///
    /// Planning owners call this immediately before returning their ordinary
    /// `Patch`; the intent is shared by its inverse and does not retain a
    /// second package snapshot.
    pub(in crate::web::package) fn with_durable_intent(mut self, intent: WebIntent) -> Self {
        self.durable_intent = Some(Arc::new(intent));
        self
    }

    /// Encode this source-bound semantic transaction in the common reversible
    /// patch envelope.
    pub fn to_durable(
        &self,
        limits: PatchLimits,
    ) -> std::result::Result<CorePatch<Reversible>, PatchError> {
        to_durable(self, limits)
    }
}

/// Apply one durable Web Extensions transaction with ordinary graph limits.
/// The caller owns the package mutation; every failure leaves it unchanged.
pub fn apply_durable<Mode>(
    package: &mut OpcPackage,
    patch: &CorePatch<Mode>,
    limits: &Limits,
) -> Result<bool> {
    apply_durable_with_limits(package, patch, limits)
}

/// Apply one durable Web Extensions transaction under caller-provided bounds.
pub fn apply_durable_with_limits<Mode>(
    package: &mut OpcPackage,
    patch: &CorePatch<Mode>,
    limits: &Limits,
) -> Result<bool> {
    let planned = plan_durable_with_limits(package, patch, limits)?;
    planned.apply(package)
}

/// Verify one durable operation and return the ordinary source-checked graph
/// [`Patch`] that will publish it. This keeps signed-package and atomicity
/// policy in the existing transaction owner.
pub fn plan_durable_with_limits<Mode>(
    package: &OpcPackage,
    patch: &CorePatch<Mode>,
    limits: &Limits,
) -> Result<Patch> {
    if patch.format() != FORMAT_NAME {
        return invalid("unsupported Web Extensions durable format".into());
    }
    if patch.operations().len() != 1 {
        return invalid("Web Extensions durable patches contain one operation".into());
    }
    let operation = &patch.operations()[0];
    if operation.target != "package" || !operation.value.is_null() {
        return invalid("invalid Web Extensions durable target or value".into());
    }
    let source = precondition(operation, "artifact_sha256")?;
    let target = precondition(operation, "target_sha256")?;
    match operation.op.as_str() {
        NOOP => {
            if operation.preconditions.len() != 2 || target != source || !patch.blobs().is_empty() {
                return invalid("invalid Web Extensions durable no-op".into());
            }
            let generated = source_noop_patch_for_package(package, limits)?;
            let maximum =
                MAX_CLOSURE_BYTES.min(limits.total_xml_bytes.max(limits.total_image_bytes));
            let actual = scope_hash_from_patch_scope(&generated, true, maximum)?;
            if actual != source {
                return Err(Error::Relationship(
                    "Web Extensions durable no-op source scope changed".into(),
                ));
            }
            Ok(generated)
        },
        EDIT => {
            if operation.preconditions.len() != 3 {
                return invalid("invalid Web Extensions durable edit".into());
            }
            let blob_id = precondition(operation, "intent_sha256")?;
            let intent = decode_intent(single_blob_by_hex(patch, blob_id)?, limits)
                .map_err(map_intent_decode_error)?;
            let generated = replay_intent(package, &intent, limits)?;
            let maximum =
                MAX_CLOSURE_BYTES.min(limits.total_xml_bytes.max(limits.total_image_bytes));
            let generated_source = scope_hash_from_patch_scope(&generated, true, maximum)?;
            let generated_target = scope_hash_from_patch_scope(&generated, false, maximum)?;
            if generated_source != source {
                return Err(Error::Relationship(
                    "Web Extensions durable typed replay target mismatch (source scope mismatch)"
                        .into(),
                ));
            }
            if generated_target != target {
                return Err(Error::Relationship(
                    "Web Extensions durable typed replay target mismatch".into(),
                ));
            }
            Ok(generated.with_durable_intent(intent))
        },
        RESTORE => {
            if operation.preconditions.len() != 3 {
                return invalid("invalid Web Extensions durable restore".into());
            }
            let blob_id = precondition(operation, "restore_sha256")?;
            let (intent_bytes, closure_bytes) =
                decode_restore(single_blob_by_hex(patch, blob_id)?, limits)?;
            let intent = decode_intent(intent_bytes, limits).map_err(map_intent_decode_error)?;
            let records = decode_closure(closure_bytes, limits)?;
            match validate_closure_source(package, &records, limits) {
                Ok(()) => {},
                Err(Error::Invalid(message)) if message.starts_with("stale Web Extensions") => {
                    return Err(Error::Relationship(
                        "Web Extensions durable restore forward replay mismatch".into(),
                    ));
                },
                Err(error) => return Err(error),
            }
            let mut candidate = package.clone();
            reconstruct_source(&mut candidate, &records, limits)?;
            let generated = replay_intent(&candidate, &intent, limits)?;
            let maximum =
                MAX_CLOSURE_BYTES.min(limits.total_xml_bytes.max(limits.total_image_bytes));
            let replayed_source = scope_hash_from_patch_scope(&generated, true, maximum)?;
            let replayed_target = scope_hash_from_patch_scope(&generated, false, maximum)?;
            let expected_closure = encode_patch_closure(&generated, true, maximum)?;
            if replayed_source != target
                || replayed_target != source
                || expected_closure.as_slice() != closure_bytes
            {
                return Err(Error::Relationship(
                    "Web Extensions durable restore forward replay mismatch".into(),
                ));
            }
            // The generated inverse carries exact source XML/relationship tokens
            // captured from the proof candidate. Apply it only after every
            // semantic check has passed, and publish atomically through a clone.
            Ok(generated.inverse().with_durable_intent(intent))
        },
        _ => invalid("unsupported Web Extensions durable operation".into()),
    }
}

fn source_noop_patch_for_package(package: &OpcPackage, limits: &Limits) -> Result<Patch> {
    let mut budget = OperationBudget::default();
    let index = super::PackageGraphIndex::build(package, limits, &mut budget)?;
    if let Some(graph) = super::existing_web_extension_graph(package, limits, &index, &mut budget)?
    {
        super::planning::source_bound_noop_patch(package, &graph, limits)
    } else {
        super::planning::empty_source_bound_noop_patch(package, limits)
    }
}

fn to_durable(
    patch: &Patch,
    limits: PatchLimits,
) -> std::result::Result<CorePatch<Reversible>, PatchError> {
    let intent = patch
        .durable_intent
        .as_deref()
        .ok_or(PatchError::InvalidText {
            field: "Web Extensions semantic intent",
        })?;
    // A no-op carries only its bounded source-scope digest.  It deliberately
    // works with zero blob allowance because no payload is retained.
    if patch.is_empty() {
        let source = scope_hash_from_patch_scope(patch, true, MAX_CLOSURE_BYTES)?;
        let target = scope_hash_from_patch_scope(patch, false, MAX_CLOSURE_BYTES)?;
        return reversible_noop(limits, &source, &target);
    }
    let operation_limit = limits.blobs();
    if operation_limit.max_blobs() == 0 {
        return Err(PatchError::BlobLimit {
            kind: BlobLimitKind::Count,
            observed: 1,
            limit: operation_limit.max_blobs(),
        });
    }
    let maximum = operation_limit
        .max_blob_bytes()
        .min(operation_limit.max_total_bytes())
        .min(MAX_CLOSURE_BYTES);
    if maximum == 0 {
        return Err(PatchError::BlobLimit {
            kind: BlobLimitKind::BlobBytes,
            observed: 1,
            limit: 0,
        });
    }
    let semantic_limits = patch
        .protection
        .as_ref()
        .map_or_else(Limits::standard, |guard| guard.limits);
    let intent_bytes = encode_intent(intent, maximum, semantic_limits)?;
    let source = scope_hash_from_patch_scope(patch, true, maximum)?;
    let target = scope_hash_from_patch_scope(patch, false, maximum)?;
    let restore_closure = encode_patch_closure(patch, !patch.durable_reversed, maximum)?;
    let restore = encode_restore(&intent_bytes, &restore_closure, maximum)?;
    let intent_id = BlobId::of(&intent_bytes);
    let restore_id = BlobId::of(&restore);
    let mut forward_blobs = BlobBundle::new(operation_limit);
    let mut reverse_blobs = BlobBundle::new(operation_limit);
    let (forward, inverse) = if patch.durable_reversed {
        forward_blobs.insert_shared(Arc::from(restore))?;
        reverse_blobs.insert_shared(Arc::from(intent_bytes))?;
        (
            operation(
                limits,
                RESTORE,
                &source,
                &target,
                Some(("restore_sha256", &restore_id)),
            )?,
            operation(
                limits,
                EDIT,
                &target,
                &source,
                Some(("intent_sha256", &intent_id)),
            )?,
        )
    } else {
        forward_blobs.insert_shared(Arc::from(intent_bytes))?;
        reverse_blobs.insert_shared(Arc::from(restore))?;
        (
            operation(
                limits,
                EDIT,
                &source,
                &target,
                Some(("intent_sha256", &intent_id)),
            )?,
            operation(
                limits,
                RESTORE,
                &target,
                &source,
                Some(("restore_sha256", &restore_id)),
            )?,
        )
    };
    CorePatch::<Reversible>::new(
        limits,
        FORMAT_NAME,
        [ReversibleOperation::new(forward, inverse)],
        forward_blobs,
        reverse_blobs,
    )
}

fn reversible_noop(
    limits: PatchLimits,
    source: &str,
    target: &str,
) -> std::result::Result<CorePatch<Reversible>, PatchError> {
    let forward = operation(limits, NOOP, source, target, None)?;
    let inverse = operation(limits, NOOP, target, source, None)?;
    CorePatch::<Reversible>::new(
        limits,
        FORMAT_NAME,
        [ReversibleOperation::new(forward, inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
}

fn operation(
    limits: PatchLimits,
    op: &str,
    source: &str,
    target: &str,
    blob: Option<(&str, &BlobId)>,
) -> std::result::Result<PatchOperation, PatchError> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert("artifact_sha256".into(), Value::String(source.into()));
    preconditions.insert("target_sha256".into(), Value::String(target.into()));
    if let Some((key, id)) = blob {
        preconditions.insert(key.into(), Value::String(id.as_hex()));
    }
    PatchOperation::new(limits, op, "package", preconditions, Value::Null)
}

fn precondition<'a>(operation: &'a PatchOperation, key: &str) -> Result<&'a str> {
    operation
        .preconditions
        .get(key)
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Invalid(format!(
                "missing Web Extensions durable precondition '{key}'"
            ))
        })
}

fn single_blob_by_hex<'a, Mode>(patch: &'a CorePatch<Mode>, id: &str) -> Result<&'a [u8]> {
    if patch.blobs().len() != 1 {
        return invalid("Web Extensions durable operation has extra blobs".into());
    }
    patch
        .blobs()
        .ids()
        .find(|candidate| candidate.as_hex() == id)
        .and_then(|candidate| patch.blobs().get(candidate))
        .ok_or_else(|| Error::Missing("Web Extensions durable semantic blob".into()))
}

struct Encoder {
    bytes: Vec<u8>,
    maximum: usize,
}

impl Encoder {
    fn new(header: &[u8], maximum: usize) -> std::result::Result<Self, PatchError> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(header.len())
            .map_err(|_| PatchError::Allocation)?;
        bytes.extend_from_slice(header);
        Ok(Self { bytes, maximum })
    }

    fn append(&mut self, value: &[u8]) -> std::result::Result<(), PatchError> {
        let size = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or(PatchError::InvalidText {
                field: "Web Extensions durable byte length",
            })?;
        if size > self.maximum {
            return Err(PatchError::InvalidText {
                field: "Web Extensions durable byte limit",
            });
        }
        self.bytes
            .try_reserve(value.len())
            .map_err(|_| PatchError::Allocation)?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn u8(&mut self, value: u8) -> std::result::Result<(), PatchError> {
        self.append(&[value])
    }
    fn bool(&mut self, value: bool) -> std::result::Result<(), PatchError> {
        self.u8(u8::from(value))
    }
    fn u32(&mut self, value: u32) -> std::result::Result<(), PatchError> {
        self.append(&value.to_le_bytes())
    }
    fn u64(&mut self, value: usize) -> std::result::Result<(), PatchError> {
        self.append(
            &u64::try_from(value)
                .map_err(|_| PatchError::InvalidText {
                    field: "Web Extensions durable integer",
                })?
                .to_le_bytes(),
        )
    }
    fn bytes(&mut self, value: &[u8]) -> std::result::Result<(), PatchError> {
        self.u64(value.len())?;
        self.append(value)
    }
    fn text(&mut self, value: &str) -> std::result::Result<(), PatchError> {
        // The enclosing encoder maximum is the caller's wire budget.  Do not
        // apply the member-name/content-type bound to XML text fields: valid
        // extension XML and external targets may be larger than 4 KiB.
        if value.len() > self.maximum {
            return Err(PatchError::InvalidText {
                field: "Web Extensions durable text",
            });
        }
        self.bytes(value.as_bytes())
    }
    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

fn encode_intent(
    intent: &WebIntent,
    maximum: usize,
    semantic_limits: Limits,
) -> std::result::Result<Vec<u8>, PatchError> {
    let mut encoder = Encoder::new(INTENT_HEADER, maximum)?;
    match intent {
        WebIntent::Put { panes, conformance } => {
            encoder.u8(0)?;
            encoder.u8(match conformance {
                Conformance::Transitional => 0,
                Conformance::Strict => 1,
            })?;
            encode_panes(&mut encoder, panes, *conformance, semantic_limits)?;
        },
        WebIntent::Remove => encoder.u8(1)?,
        WebIntent::CustomFunctions(edit) => {
            encoder.u8(2)?;
            encoder.u8(match edit.mode {
                GraphMode::ExistingOnly => 0,
                GraphMode::EnsureGraph => 1,
                GraphMode::ReplaceManagedGraph => 2,
            })?;
            encoder.u64(edit.additions.len())?;
            for pane in &edit.additions {
                encode_pane(
                    &mut encoder,
                    pane,
                    Conformance::Transitional,
                    semantic_limits,
                )?;
            }
            encoder.u64(edit.operations.len())?;
            for (selector, operation) in &edit.operations {
                encode_selector(&mut encoder, selector)?;
                encode_operation(&mut encoder, operation)?;
            }
        },
    }
    Ok(encoder.finish())
}

fn encode_panes(
    encoder: &mut Encoder,
    panes: &Panes,
    conformance: Conformance,
    semantic_limits: Limits,
) -> std::result::Result<(), PatchError> {
    let xml = write_panes_with(panes, conformance, &semantic_limits).map_err(|_| {
        PatchError::InvalidText {
            field: "Web Extensions panes",
        }
    })?;
    encoder.bytes(&xml)?;
    encoder.u64(panes.len())?;
    for pane in panes.iter() {
        encode_pane(encoder, pane, conformance, semantic_limits)?;
    }
    Ok(())
}

fn encode_pane(
    encoder: &mut Encoder,
    pane: &Pane,
    conformance: Conformance,
    semantic_limits: Limits,
) -> std::result::Result<(), PatchError> {
    encoder.text(pane.dock_state())?;
    encoder.bool(pane.visible())?;
    encoder.append(&pane.pane_width().to_bits().to_le_bytes())?;
    encoder.u32(pane.row())?;
    encoder.bool(pane.locked())?;
    encoder.text(&pane.relationship_id)?;
    let xml = write_add_in_with(&pane.add_in, conformance, &semantic_limits).map_err(|_| {
        PatchError::InvalidText {
            field: "Web Extensions add-in",
        }
    })?;
    encoder.bytes(&xml)?;
    encoder.u64(pane.snapshot_resources.len())?;
    for resource in &pane.snapshot_resources {
        encoder.text(&resource.relationship_id)?;
        match &resource.target {
            SnapshotTarget::Internal {
                part_name,
                content_type,
                data,
            } => {
                encoder.u8(0)?;
                encoder.text(part_name.as_str())?;
                encoder.text(content_type)?;
                encoder.bytes(data)?;
            },
            SnapshotTarget::External { target } => {
                encoder.u8(1)?;
                encoder.text(target)?;
            },
        }
    }
    encoder.bool(pane.extension_list.is_some())?;
    if let Some(extension) = &pane.extension_list {
        encoder.text(extension.xml())?;
    }
    Ok(())
}

fn encode_selector(
    encoder: &mut Encoder,
    selector: &OwnerSelector,
) -> std::result::Result<(), PatchError> {
    match selector {
        OwnerSelector::AddInId(id) => {
            encoder.u8(0)?;
            encoder.text(id)?;
        },
        OwnerSelector::PaneIndex(index) => {
            encoder.u8(1)?;
            encoder.u64(*index)?;
        },
    }
    Ok(())
}

fn encode_operation(
    encoder: &mut Encoder,
    operation: &CustomFunctionOperation,
) -> std::result::Result<(), PatchError> {
    match operation {
        CustomFunctionOperation::Set(value) => {
            encoder.u8(0)?;
            encode_custom_functions(encoder, value.as_ref())?;
        },
        CustomFunctionOperation::SetContains(value) => {
            encoder.u8(1)?;
            encode_contains(encoder, value.as_ref())?;
        },
        CustomFunctionOperation::SetBackground(value) => {
            encoder.u8(2)?;
            encode_background(encoder, value.as_ref())?;
        },
        CustomFunctionOperation::SetList(value) => {
            encoder.u8(3)?;
            encode_list(encoder, value.as_ref())?;
        },
        CustomFunctionOperation::InsertId { index, id } => {
            encoder.u8(4)?;
            encoder.u64(*index)?;
            encoder.text(id)?;
        },
        CustomFunctionOperation::ReplaceId { index, id } => {
            encoder.u8(5)?;
            encoder.u64(*index)?;
            encoder.text(id)?;
        },
        CustomFunctionOperation::RemoveId { index } => {
            encoder.u8(6)?;
            encoder.u64(*index)?;
        },
    }
    Ok(())
}

fn encode_custom_functions(
    encoder: &mut Encoder,
    value: Option<&CustomFunctions>,
) -> std::result::Result<(), PatchError> {
    encoder.bool(value.is_some())?;
    let Some(value) = value else { return Ok(()) };
    encode_contains(encoder, value.contains_custom_functions())?;
    encode_background(encoder, value.background_app_data())?;
    encode_list(encoder, value.custom_function_list())
}

fn encode_contains(
    encoder: &mut Encoder,
    value: Option<&ContainsCustomFunctions>,
) -> std::result::Result<(), PatchError> {
    encoder.bool(value.is_some())?;
    if let Some(value) = value {
        encoder.bool(value.explicit_value().is_some())?;
        encoder.bool(value.value())?;
    }
    Ok(())
}

fn encode_background(
    encoder: &mut Encoder,
    value: Option<&BackgroundAppData>,
) -> std::result::Result<(), PatchError> {
    encoder.bool(value.is_some())?;
    if let Some(value) = value {
        encoder.append(&value.state().to_le_bytes())?;
        encoder.text(value.runtime_id())?;
    }
    Ok(())
}

fn encode_list(
    encoder: &mut Encoder,
    value: Option<&CustomFunctionList>,
) -> std::result::Result<(), PatchError> {
    encoder.bool(value.is_some())?;
    if let Some(value) = value {
        encoder.u64(value.ids().len())?;
        for id in value.ids() {
            encoder.text(id)?;
        }
    }
    Ok(())
}

fn encode_restore(
    intent: &[u8],
    closure: &[u8],
    maximum: usize,
) -> std::result::Result<Vec<u8>, PatchError> {
    let mut encoder = Encoder::new(RESTORE_HEADER, maximum)?;
    encoder.bytes(intent)?;
    encoder.bytes(closure)?;
    Ok(encoder.finish())
}

struct Decoder<'a> {
    bytes: &'a [u8],
    offset: usize,
    string_total: usize,
    string_maximum: usize,
}

/// Shared semantic accounting for one decoded intent. The ordinary model
/// parsers each accept an [`OperationBudget`]; keeping one instance here makes
/// their XML and retained-string work cumulative across every pane. XML node
/// and depth limits remain per-document `Limits` because the parser does not
/// retain a cross-document node inventory. Snapshot images are inert bytes,
/// so they use the parallel image ledger.
#[derive(Debug, Default)]
struct IntentDecodeBudget {
    semantic: OperationBudget,
    image_bytes: usize,
    images: Vec<DecodedImage>,
    image_indices: HashMap<Vec<u8>, usize>,
}

#[derive(Debug)]
struct DecodedImage {
    content_type: String,
    data: Arc<Vec<u8>>,
}

impl IntentDecodeBudget {
    fn preflight_xml(&self, bytes: usize, limits: &Limits) -> Result<()> {
        let total = self
            .semantic
            .xml_bytes
            .checked_add(bytes)
            .ok_or(Error::Limit {
                resource: "aggregate web extension XML bytes",
                max: limits.total_xml_bytes,
                actual: usize::MAX,
            })?;
        if total > limits.total_xml_bytes {
            return Err(Error::Limit {
                resource: "aggregate web extension XML bytes",
                max: limits.total_xml_bytes,
                actual: total,
            });
        }
        Ok(())
    }

    fn preflight_image(&self, bytes: usize, limits: &Limits) -> Result<()> {
        let total = self.image_bytes.checked_add(bytes).ok_or(Error::Limit {
            resource: "aggregate web extension image bytes",
            max: limits.total_image_bytes,
            actual: usize::MAX,
        })?;
        if total > limits.total_image_bytes {
            return Err(Error::Limit {
                resource: "aggregate web extension image bytes",
                max: limits.total_image_bytes,
                actual: total,
            });
        }
        Ok(())
    }

    fn charge_extension_xml(&mut self, bytes: usize, limits: &Limits) -> Result<()> {
        self.preflight_xml(bytes, limits)?;
        self.semantic.charge_xml(bytes, limits)?;
        self.semantic.charge_strings(bytes, limits)
    }
}

impl<'a> Decoder<'a> {
    fn new(bytes: &'a [u8], header: &[u8], string_maximum: usize) -> Result<Self> {
        if !bytes.starts_with(header) {
            return invalid("invalid Web Extensions durable semantic header".into());
        }
        Ok(Self {
            bytes,
            offset: header.len(),
            string_total: 0,
            string_maximum,
        })
    }

    fn take(&mut self, length: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(length)
            .ok_or_else(|| Error::Invalid("Web Extensions durable offset overflow".into()))?;
        let value = self
            .bytes
            .get(self.offset..end)
            .ok_or_else(|| Error::Invalid("truncated Web Extensions durable blob".into()))?;
        self.offset = end;
        Ok(value)
    }

    fn u8(&mut self) -> Result<u8> {
        Ok(self.take(1)?[0])
    }

    fn bool(&mut self, field: &'static str) -> Result<bool> {
        match self.u8()? {
            0 => Ok(false),
            1 => Ok(true),
            _ => invalid(format!(
                "invalid Web Extensions durable boolean for {field}"
            )),
        }
    }

    fn u64(&mut self, field: &'static str, maximum: usize) -> Result<usize> {
        let bytes: [u8; 8] = self
            .take(8)?
            .try_into()
            .map_err(|_| Error::Invalid("invalid Web Extensions durable integer".into()))?;
        let value = usize::try_from(u64::from_le_bytes(bytes)).map_err(|_| {
            Error::Invalid("Web Extensions durable integer exceeds this platform".into())
        })?;
        if value > maximum {
            return Err(Error::Limit {
                resource: field,
                max: maximum,
                actual: value,
            });
        }
        Ok(value)
    }

    fn u32(&mut self, field: &'static str) -> Result<u32> {
        let bytes: [u8; 4] = self
            .take(4)?
            .try_into()
            .map_err(|_| Error::Invalid(format!("invalid {field}")))?;
        Ok(u32::from_le_bytes(bytes))
    }

    fn i32(&mut self) -> Result<i32> {
        Ok(i32::from_le_bytes(self.take(4)?.try_into().map_err(
            |_| Error::Invalid("invalid Web Extensions durable i32".into()),
        )?))
    }

    fn f64(&mut self) -> Result<f64> {
        Ok(f64::from_bits(u64::from_le_bytes(
            self.take(8)?
                .try_into()
                .map_err(|_| Error::Invalid("invalid Web Extensions durable f64".into()))?,
        )))
    }

    fn bytes_slice(&mut self, field: &'static str, maximum: usize) -> Result<&'a [u8]> {
        let length = self.u64(field, maximum)?;
        self.take(length)
    }

    fn copy_source(source: &[u8], resource: &'static str) -> Result<Vec<u8>> {
        let mut output = Vec::new();
        output
            .try_reserve_exact(source.len())
            .map_err(|source| Error::Allocation { resource, source })?;
        output.extend_from_slice(source);
        Ok(output)
    }

    fn bytes_with_total(
        &mut self,
        field: &'static str,
        maximum: usize,
        total: &mut usize,
        total_maximum: usize,
        resource: &'static str,
    ) -> Result<Vec<u8>> {
        let source = self.bytes_slice(field, maximum)?;
        let next = total.checked_add(source.len()).ok_or(Error::Limit {
            resource,
            max: total_maximum,
            actual: usize::MAX,
        })?;
        if next > total_maximum {
            return Err(Error::Limit {
                resource,
                max: total_maximum,
                actual: next,
            });
        }
        let output = Self::copy_source(source, "Web Extensions durable bytes")?;
        *total = next;
        Ok(output)
    }

    fn synchronize_string_budget(
        &mut self,
        budget: &mut OperationBudget,
        limits: &Limits,
    ) -> Result<()> {
        if self.string_total > budget.string_bytes {
            budget.string_bytes = self.string_total;
        } else {
            self.string_total = budget.string_bytes;
        }
        if budget.string_bytes > limits.total_string_bytes {
            return Err(Error::Limit {
                resource: "retained web extension string bytes",
                max: limits.total_string_bytes,
                actual: budget.string_bytes,
            });
        }
        Ok(())
    }

    fn text(&mut self, field: &'static str, maximum: usize) -> Result<String> {
        let bytes = self.bytes_slice(field, maximum)?;
        let text = std::str::from_utf8(bytes)
            .map_err(|_| Error::Invalid("Web Extensions durable text is not UTF-8".into()))?;
        if text.contains('\0') {
            return invalid("Web Extensions durable text contains NUL".into());
        }
        let next = self
            .string_total
            .checked_add(text.len())
            .ok_or(Error::Limit {
                resource: "Web Extensions durable decoded strings",
                max: self.string_maximum,
                actual: usize::MAX,
            })?;
        if next > self.string_maximum {
            return Err(Error::Limit {
                resource: "Web Extensions durable decoded strings",
                max: self.string_maximum,
                actual: next,
            });
        }
        let mut output = String::new();
        output
            .try_reserve_exact(text.len())
            .map_err(|source| Error::Allocation {
                resource: "Web Extensions durable decoded strings",
                source,
            })?;
        output.push_str(text);
        self.string_total = next;
        Ok(output)
    }

    fn finished(&self) -> bool {
        self.offset == self.bytes.len()
    }
    fn remaining(&self) -> usize {
        self.bytes.len().saturating_sub(self.offset)
    }
}

fn intent_xml_for_parser(
    decoder: &mut Decoder<'_>,
    field: &'static str,
    limits: &Limits,
    budget: &mut IntentDecodeBudget,
) -> Result<(Vec<u8>, Limits)> {
    decoder.synchronize_string_budget(&mut budget.semantic, limits)?;
    let source = decoder.bytes_slice(field, limits.xml_bytes)?;
    budget.preflight_xml(source.len(), limits)?;
    let parser_limits = parser_limits_for_budget(limits, &budget.semantic, source.len())?;
    Ok((
        Decoder::copy_source(source, "Web Extensions durable XML bytes")?,
        parser_limits,
    ))
}

fn intent_extension_xml(
    decoder: &mut Decoder<'_>,
    field: &'static str,
    limits: &Limits,
    budget: &mut IntentDecodeBudget,
) -> Result<(Vec<u8>, Limits)> {
    decoder.synchronize_string_budget(&mut budget.semantic, limits)?;
    let source = decoder.bytes_slice(field, limits.xml_bytes)?;
    budget.charge_extension_xml(source.len(), limits)?;
    decoder.synchronize_string_budget(&mut budget.semantic, limits)?;
    let parser_limits = parser_limits_for_budget(limits, &budget.semantic, source.len())?;
    Ok((
        Decoder::copy_source(source, "Web Extensions durable XML bytes")?,
        parser_limits,
    ))
}

fn parser_limits_for_budget(
    limits: &Limits,
    budget: &OperationBudget,
    source_len: usize,
) -> Result<Limits> {
    let remaining = limits
        .total_string_bytes
        .checked_sub(budget.string_bytes)
        .ok_or(Error::Limit {
            resource: "retained web extension string bytes",
            max: limits.total_string_bytes,
            actual: budget.string_bytes,
        })?;
    let next = budget
        .string_bytes
        .checked_add(source_len)
        .ok_or(Error::Limit {
            resource: "retained web extension string bytes",
            max: limits.total_string_bytes,
            actual: usize::MAX,
        })?;
    if next > limits.total_string_bytes {
        return Err(Error::Limit {
            resource: "retained web extension string bytes",
            max: limits.total_string_bytes,
            actual: next,
        });
    }
    let mut parser_limits = *limits;
    // The parser retains its processed XML alongside decoded names and values.
    // Reserve the raw field size before parsing so a large second document
    // cannot be materialized and only then fail the shared string ledger.
    parser_limits.xml_bytes = parser_limits.xml_bytes.min(source_len).min(remaining);
    parser_limits.string_bytes = parser_limits
        .string_bytes
        .min(remaining.saturating_sub(source_len));
    Ok(parser_limits)
}

fn intent_image_shared(
    decoder: &mut Decoder<'_>,
    field: &'static str,
    limits: &Limits,
    part_name: &PackURI,
    content_type: &str,
    budget: &mut IntentDecodeBudget,
) -> Result<Arc<Vec<u8>>> {
    let source = decoder.bytes_slice(field, limits.image_bytes)?;
    let folded_name = ascii_folded_name(part_name.as_str())?;
    if let Some(&index) = budget.image_indices.get(&folded_name) {
        let existing = budget.images.get(index).ok_or_else(|| {
            Error::Invalid("Web Extensions durable image ledger index is invalid".into())
        })?;
        if existing.content_type == content_type && existing.data.as_slice() == source {
            return Ok(Arc::clone(&existing.data));
        }
        return invalid(format!(
            "conflicting durable image payload for part '{}'",
            part_name.as_str()
        ));
    }
    decoder.synchronize_string_budget(&mut budget.semantic, limits)?;
    let ledger_metadata =
        folded_name
            .len()
            .checked_add(content_type.len())
            .ok_or(Error::Limit {
                resource: "retained web extension string bytes",
                max: limits.total_string_bytes,
                actual: usize::MAX,
            })?;
    budget.semantic.charge_strings(ledger_metadata, limits)?;
    decoder.synchronize_string_budget(&mut budget.semantic, limits)?;
    budget.preflight_image(source.len(), limits)?;
    let mut ledger_content_type = String::new();
    ledger_content_type
        .try_reserve_exact(content_type.len())
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions durable image ledger",
            source,
        })?;
    ledger_content_type.push_str(content_type);
    let output = Decoder::copy_source(source, "Web Extensions durable image bytes")?;
    budget.image_bytes = budget
        .image_bytes
        .checked_add(source.len())
        .ok_or(Error::Limit {
            resource: "aggregate web extension image bytes",
            max: limits.total_image_bytes,
            actual: usize::MAX,
        })?;
    let data = Arc::new(output);
    budget.images.push(DecodedImage {
        content_type: ledger_content_type,
        data: Arc::clone(&data),
    });
    budget
        .image_indices
        .insert(folded_name, budget.images.len() - 1);
    Ok(data)
}

fn decode_intent(bytes: &[u8], limits: &Limits) -> Result<WebIntent> {
    let maximum = MAX_INTENT_BYTES;
    if bytes.len() > maximum {
        return Err(Error::Limit {
            resource: "Web Extensions durable intent bytes",
            max: maximum,
            actual: bytes.len(),
        });
    }
    let mut decoder = Decoder::new(bytes, INTENT_HEADER, limits.total_string_bytes)?;
    let mut budget = IntentDecodeBudget::default();
    let intent = match decoder.u8()? {
        0 => {
            let conformance = match decoder.u8()? {
                0 => Conformance::Transitional,
                1 => Conformance::Strict,
                _ => return invalid("unknown Web Extensions conformance code".into()),
            };
            WebIntent::Put {
                panes: decode_panes(&mut decoder, limits, &mut budget)?,
                conformance,
            }
        },
        1 => WebIntent::Remove,
        2 => {
            let mode = match decoder.u8()? {
                0 => GraphMode::ExistingOnly,
                1 => GraphMode::EnsureGraph,
                2 => GraphMode::ReplaceManagedGraph,
                _ => return invalid("unknown Web Extensions graph mode".into()),
            };
            let addition_count = decoder.u64("Web Extensions durable pane count", limits.items)?;
            const MIN_PANE_WIRE_BYTES: usize = 47;
            if addition_count > decoder.remaining() / MIN_PANE_WIRE_BYTES {
                return invalid("truncated Web Extensions durable panes".into());
            }
            let retained = addition_count
                .checked_mul(size_of::<Pane>())
                .ok_or_else(|| Error::Limit {
                    resource: "Web Extensions durable authored panes",
                    max: limits.total_xml_bytes.max(limits.total_image_bytes),
                    actual: usize::MAX,
                })?;
            let staging_maximum = limits
                .total_xml_bytes
                .max(limits.total_image_bytes)
                .max(limits.total_string_bytes);
            if retained > staging_maximum {
                return Err(Error::Limit {
                    resource: "Web Extensions durable authored panes",
                    max: staging_maximum,
                    actual: retained,
                });
            }
            let mut additions = Vec::new();
            additions
                .try_reserve_exact(addition_count)
                .map_err(|source| Error::Allocation {
                    resource: "Web Extensions durable authored panes",
                    source,
                })?;
            for _ in 0..addition_count {
                additions.push(decode_pane(&mut decoder, limits, &mut budget)?);
            }
            let operation_count =
                decoder.u64("Web Extensions durable operation count", limits.items)?;
            if operation_count > MAX_INTENT_OPERATIONS {
                return Err(Error::Limit {
                    resource: "Web Extensions durable operations",
                    max: MAX_INTENT_OPERATIONS,
                    actual: operation_count,
                });
            }
            const MIN_OPERATION_WIRE_BYTES: usize = 1 + 8 + 1 + 1;
            if operation_count > decoder.remaining() / MIN_OPERATION_WIRE_BYTES {
                return invalid("truncated Web Extensions durable operations".into());
            }
            let retained = operation_count
                .checked_mul(size_of::<(OwnerSelector, CustomFunctionOperation)>())
                .ok_or_else(|| Error::Limit {
                    resource: "Web Extensions durable operations",
                    max: limits.total_string_bytes.max(limits.total_xml_bytes),
                    actual: usize::MAX,
                })?;
            let staging_maximum = limits
                .total_xml_bytes
                .max(limits.total_image_bytes)
                .max(limits.total_string_bytes);
            if retained > staging_maximum {
                return Err(Error::Limit {
                    resource: "Web Extensions durable operations",
                    max: staging_maximum,
                    actual: retained,
                });
            }
            let mut operations = Vec::new();
            operations
                .try_reserve_exact(operation_count)
                .map_err(|source| Error::Allocation {
                    resource: "Web Extensions durable operations",
                    source,
                })?;
            for _ in 0..operation_count {
                operations.push((
                    decode_selector(&mut decoder, limits)?,
                    decode_operation(&mut decoder, limits)?,
                ));
            }
            WebIntent::CustomFunctions(CustomFunctionEdit {
                mode,
                operations,
                additions,
            })
        },
        _ => return invalid("unknown Web Extensions durable intent".into()),
    };
    if !decoder.finished() {
        return invalid("trailing Web Extensions durable intent bytes".into());
    }
    Ok(intent)
}

fn map_intent_decode_error(error: Error) -> Error {
    if matches!(
        &error,
        Error::Invalid(message) if message.contains("text is not permitted")
    ) {
        Error::Relationship("Web Extensions durable typed replay target mismatch".into())
    } else {
        error
    }
}

fn decode_panes(
    decoder: &mut Decoder<'_>,
    limits: &Limits,
    budget: &mut IntentDecodeBudget,
) -> Result<Panes> {
    let (xml, parser_limits) =
        intent_xml_for_parser(decoder, "Web Extensions task-pane XML", limits, budget)?;
    let parsed = parse_panes_with_budget(&xml, &parser_limits, &mut budget.semantic)?;
    decoder.synchronize_string_budget(&mut budget.semantic, limits)?;
    let count = decoder.u64("Web Extensions durable pane count", limits.items)?;
    if count != parsed.len() {
        return invalid("Web Extensions durable pane count disagrees with XML".into());
    }
    const MIN_PANE_WIRE_BYTES: usize = 47;
    if count > decoder.remaining() / MIN_PANE_WIRE_BYTES {
        return invalid("truncated Web Extensions durable panes".into());
    }
    let retained = count
        .checked_mul(size_of::<Pane>())
        .ok_or_else(|| Error::Limit {
            resource: "Web Extensions durable panes",
            max: limits.total_xml_bytes.max(limits.total_image_bytes),
            actual: usize::MAX,
        })?;
    let staging_maximum = limits
        .total_xml_bytes
        .max(limits.total_image_bytes)
        .max(limits.total_string_bytes);
    if retained > staging_maximum {
        return Err(Error::Limit {
            resource: "Web Extensions durable panes",
            max: staging_maximum,
            actual: retained,
        });
    }
    let mut panes = Vec::new();
    panes
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions durable panes",
            source,
        })?;
    for parsed_pane in parsed {
        // The task-pane XML carries the ordered pane attributes and relationship
        // ID; the following payload carries its add-in and snapshot closure.
        let payload = decode_pane(decoder, limits, budget)?;
        if payload.dock_state != parsed_pane.dock_state
            || payload.visible != parsed_pane.visible
            || payload.width != parsed_pane.width
            || payload.row != parsed_pane.row
            || payload.locked != parsed_pane.locked
            || payload.relationship_id != parsed_pane.relationship_id
            || payload.extension_list != parsed_pane.extension_list
        {
            return invalid(
                "Web Extensions durable pane payload disagrees with task-pane XML".into(),
            );
        }
        panes.push(payload);
    }
    Ok(Panes { panes })
}

fn decode_pane(
    decoder: &mut Decoder<'_>,
    limits: &Limits,
    budget: &mut IntentDecodeBudget,
) -> Result<Pane> {
    let dock_state = Dock::parse(&decoder.text("Web Extensions dock state", limits.xml_bytes)?)?;
    let visible = decoder.bool("Web Extensions pane visibility")?;
    let width = decoder.f64()?;
    if !width.is_finite() || width <= 0.0 {
        return invalid("Web Extensions durable pane width is invalid".into());
    }
    let row = decoder.u32("Web Extensions pane row")?;
    let locked = decoder.bool("Web Extensions pane locked state")?;
    let relationship_id = decoder.text("Web Extensions pane relationship ID", limits.xml_bytes)?;
    let (add_in_xml, parser_limits) =
        intent_xml_for_parser(decoder, "Web Extensions add-in XML", limits, budget)?;
    let add_in = parse_add_in_with_budget(&add_in_xml, &parser_limits, &mut budget.semantic)?;
    decoder.synchronize_string_budget(&mut budget.semantic, limits)?;
    let resource_count = decoder.u64("Web Extensions snapshot resource count", limits.items)?;
    const MIN_RESOURCE_WIRE_BYTES: usize = 1 + 8;
    if resource_count > decoder.remaining() / MIN_RESOURCE_WIRE_BYTES {
        return invalid("truncated Web Extensions durable snapshot resources".into());
    }
    let retained = resource_count
        .checked_mul(size_of::<SnapshotResource>())
        .ok_or(Error::Limit {
            resource: "Web Extensions durable snapshot resources",
            max: limits
                .total_xml_bytes
                .max(limits.total_image_bytes)
                .max(limits.total_string_bytes),
            actual: usize::MAX,
        })?;
    let ledger_retained = resource_count
        .checked_mul(size_of::<DecodedImage>())
        .ok_or(Error::Limit {
            resource: "Web Extensions durable image ledger",
            max: limits
                .total_xml_bytes
                .max(limits.total_image_bytes)
                .max(limits.total_string_bytes),
            actual: usize::MAX,
        })?;
    let index_retained = resource_count
        .checked_mul(size_of::<(Vec<u8>, usize)>())
        .ok_or(Error::Limit {
            resource: "Web Extensions durable image index",
            max: limits
                .total_xml_bytes
                .max(limits.total_image_bytes)
                .max(limits.total_string_bytes),
            actual: usize::MAX,
        })?;
    let staging = retained
        .checked_add(ledger_retained)
        .and_then(|value| value.checked_add(index_retained))
        .ok_or(Error::Limit {
            resource: "Web Extensions durable snapshot staging",
            max: limits
                .total_xml_bytes
                .max(limits.total_image_bytes)
                .max(limits.total_string_bytes),
            actual: usize::MAX,
        })?;
    let staging_maximum = limits
        .total_xml_bytes
        .max(limits.total_image_bytes)
        .max(limits.total_string_bytes);
    if staging > staging_maximum {
        return Err(Error::Limit {
            resource: "Web Extensions durable snapshot staging",
            max: staging_maximum,
            actual: staging,
        });
    }
    let mut resources = Vec::new();
    resources
        .try_reserve_exact(resource_count)
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions durable snapshot resources",
            source,
        })?;
    budget
        .images
        .try_reserve(resource_count)
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions durable image ledger",
            source,
        })?;
    budget
        .image_indices
        .try_reserve(resource_count)
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions durable image index",
            source,
        })?;
    for _ in 0..resource_count {
        resources.push(decode_resource(decoder, limits, budget)?);
    }
    let extension_list = if decoder.bool("Web Extensions pane extension-list presence")? {
        let (xml, parser_limits) = intent_extension_xml(
            decoder,
            "Web Extensions pane extension-list XML",
            limits,
            budget,
        )?;
        let document = parse_xml_owned(xml, &parser_limits)?;
        let root = document.root()?;
        require_name(root, TASK_PANES_NAMESPACE, "extLst")?;
        Some(ExtList::from_node_with_limits(
            root,
            &document,
            &parser_limits,
        )?)
    } else {
        None
    };
    Ok(Pane {
        dock_state,
        visible,
        width,
        row,
        locked,
        relationship_id,
        add_in,
        snapshot_resources: resources,
        extension_list,
    })
}

fn decode_resource(
    decoder: &mut Decoder<'_>,
    limits: &Limits,
    budget: &mut IntentDecodeBudget,
) -> Result<SnapshotResource> {
    let relationship_id =
        decoder.text("Web Extensions snapshot relationship ID", limits.xml_bytes)?;
    let target = match decoder.u8()? {
        0 => {
            let name = decoder.text("Web Extensions snapshot part name", MAX_MEMBER_NAME)?;
            let content_type =
                decoder.text("Web Extensions snapshot content type", MAX_CONTENT_TYPE)?;
            let part_name = PackURI::new(name).map_err(Error::Uri)?;
            let data = intent_image_shared(
                decoder,
                "Web Extensions snapshot bytes",
                limits,
                &part_name,
                &content_type,
                budget,
            )?;
            SnapshotTarget::Internal {
                part_name,
                content_type,
                data,
            }
        },
        1 => SnapshotTarget::External {
            target: decoder.text("Web Extensions external snapshot target", limits.xml_bytes)?,
        },
        _ => return invalid("unknown Web Extensions snapshot target".into()),
    };
    Ok(SnapshotResource {
        relationship_id,
        target,
    })
}

fn decode_selector(decoder: &mut Decoder<'_>, limits: &Limits) -> Result<OwnerSelector> {
    match decoder.u8()? {
        0 => Ok(OwnerSelector::AddInId(
            decoder.text("Web Extensions owner ID", limits.xml_bytes)?,
        )),
        1 => Ok(OwnerSelector::PaneIndex(
            decoder.u64("Web Extensions pane index", limits.items)?,
        )),
        _ => invalid("unknown Web Extensions owner selector".into()),
    }
}

fn decode_operation(decoder: &mut Decoder<'_>, limits: &Limits) -> Result<CustomFunctionOperation> {
    Ok(match decoder.u8()? {
        0 => CustomFunctionOperation::Set(decode_custom_functions(decoder, limits)?),
        1 => CustomFunctionOperation::SetContains(decode_contains(decoder)?),
        2 => CustomFunctionOperation::SetBackground(decode_background(decoder, limits)?),
        3 => CustomFunctionOperation::SetList(decode_list(decoder, limits)?),
        4 => CustomFunctionOperation::InsertId {
            index: decoder.u64("Web Extensions custom-function index", limits.items)?,
            id: decoder.text("Web Extensions custom-function ID", limits.xml_bytes)?,
        },
        5 => CustomFunctionOperation::ReplaceId {
            index: decoder.u64("Web Extensions custom-function index", limits.items)?,
            id: decoder.text("Web Extensions custom-function ID", limits.xml_bytes)?,
        },
        6 => CustomFunctionOperation::RemoveId {
            index: decoder.u64("Web Extensions custom-function index", limits.items)?,
        },
        _ => return invalid("unknown Web Extensions custom-function operation".into()),
    })
}

fn decode_custom_functions(
    decoder: &mut Decoder<'_>,
    limits: &Limits,
) -> Result<Option<CustomFunctions>> {
    if !decoder.bool("custom-functions presence")? {
        return Ok(None);
    }
    let mut result = CustomFunctions::new();
    result
        .set_contains_custom_functions(decode_contains(decoder)?)
        .set_background_app_data(decode_background(decoder, limits)?)
        .set_custom_function_list(decode_list(decoder, limits)?);
    Ok(Some(result))
}

fn decode_contains(decoder: &mut Decoder<'_>) -> Result<Option<ContainsCustomFunctions>> {
    if !decoder.bool("contains-custom-functions presence")? {
        return Ok(None);
    }
    let explicit = decoder.bool("contains-custom-functions explicit value")?;
    let value = decoder.bool("contains-custom-functions value")?;
    Ok(Some(ContainsCustomFunctions::new(
        explicit.then_some(value),
    )))
}

fn decode_background(
    decoder: &mut Decoder<'_>,
    limits: &Limits,
) -> Result<Option<BackgroundAppData>> {
    if !decoder.bool("background-app-data presence")? {
        return Ok(None);
    }
    Ok(Some(BackgroundAppData::new(
        decoder.i32()?,
        decoder.text("background runtime ID", limits.xml_bytes)?,
    )?))
}

fn decode_list(decoder: &mut Decoder<'_>, limits: &Limits) -> Result<Option<CustomFunctionList>> {
    if !decoder.bool("custom-function-list presence")? {
        return Ok(None);
    }
    let count = decoder.u64("custom-function ID count", limits.items)?;
    if count > decoder.remaining() / size_of::<u64>() {
        return invalid("truncated Web Extensions custom-function IDs".into());
    }
    let retained = count.checked_mul(size_of::<String>()).ok_or(Error::Limit {
        resource: "Web Extensions custom-function IDs",
        max: limits.total_string_bytes,
        actual: usize::MAX,
    })?;
    let staging_maximum = limits
        .total_xml_bytes
        .max(limits.total_image_bytes)
        .max(limits.total_string_bytes);
    if retained > staging_maximum {
        return Err(Error::Limit {
            resource: "Web Extensions custom-function IDs",
            max: staging_maximum,
            actual: retained,
        });
    }
    let mut list = CustomFunctionList::new();
    for _ in 0..count {
        list.push_id(decoder.text("custom-function ID", limits.xml_bytes)?)?;
    }
    Ok(Some(list))
}

fn replay_intent(package: &OpcPackage, intent: &WebIntent, limits: &Limits) -> Result<Patch> {
    match intent {
        WebIntent::Put { panes, conformance } => {
            super::plan_put_with(package, panes.clone(), *conformance, limits)
        },
        WebIntent::Remove => super::plan_remove_with(package, limits),
        WebIntent::CustomFunctions(edit) => replay_custom_functions(package, edit, limits),
    }
}

fn replay_custom_functions(
    package: &OpcPackage,
    edit: &CustomFunctionEdit,
    limits: &Limits,
) -> Result<Patch> {
    if edit.mode == GraphMode::ReplaceManagedGraph {
        return invalid("ReplaceManagedGraph requires a complete WebIntent::Put graph".into());
    }
    let existing = super::load_with(package, limits)?;
    let mut panes = match existing {
        Some(panes) => panes,
        None if edit.mode == GraphMode::ExistingOnly || edit.additions.is_empty() => {
            return invalid("custom-function edit requires an existing task-pane graph".into());
        },
        None => Panes::new(),
    };
    if edit.mode == GraphMode::ExistingOnly && !edit.additions.is_empty() {
        return invalid("existing-only custom-function edit contains authored panes".into());
    }
    for pane in edit.additions.iter().cloned() {
        panes.push(pane)?;
    }
    for (selector, operation) in &edit.operations {
        let selector = match selector {
            OwnerSelector::AddInId(id) => Selector::Id(id.as_str()),
            OwnerSelector::PaneIndex(index) => Selector::Index(*index),
        };
        if !panes.edit(selector, |pane| {
            apply_operation(pane.add_in_mut(), operation, limits)
        })? {
            return Err(Error::Missing(
                "custom-function owner selector not found".into(),
            ));
        }
    }
    super::plan_put_with(package, panes, infer_conformance(package, limits)?, limits)
}

fn infer_conformance(package: &OpcPackage, limits: &Limits) -> Result<Conformance> {
    let mut budget = OperationBudget::default();
    let index = super::PackageGraphIndex::build(package, limits, &mut budget)?;
    let Some(graph) = super::existing_web_extension_graph(package, limits, &index, &mut budget)?
    else {
        return Ok(Conformance::Transitional);
    };
    if super::planning::source_matches_conformance(package, &graph, Conformance::Strict, limits)? {
        Ok(Conformance::Strict)
    } else {
        Ok(Conformance::Transitional)
    }
}

fn apply_operation(
    add_in: &mut AddIn,
    operation: &CustomFunctionOperation,
    limits: &Limits,
) -> Result<()> {
    match operation {
        CustomFunctionOperation::Set(value) => {
            add_in.set_custom_functions_with_limits(value.clone(), limits)?;
        },
        CustomFunctionOperation::SetContains(value) => {
            let mut current = add_in.custom_functions().cloned().unwrap_or_default();
            current.set_contains_custom_functions(value.clone());
            add_in.set_custom_functions_with_limits(Some(current), limits)?;
        },
        CustomFunctionOperation::SetBackground(value) => {
            let mut current = add_in.custom_functions().cloned().unwrap_or_default();
            current.set_background_app_data(value.clone());
            add_in.set_custom_functions_with_limits(Some(current), limits)?;
        },
        CustomFunctionOperation::SetList(value) => {
            let mut current = add_in.custom_functions().cloned().unwrap_or_default();
            current.set_custom_function_list(value.clone());
            add_in.set_custom_functions_with_limits(Some(current), limits)?;
        },
        CustomFunctionOperation::InsertId { index, id } => {
            let mut current = add_in.custom_functions().cloned().unwrap_or_default();
            let mut ids = current
                .custom_function_list()
                .map_or_else(Vec::new, |list| list.ids().to_vec());
            if *index > ids.len() {
                return Err(Error::Invalid(
                    "custom-function insertion index is out of range".into(),
                ));
            }
            ids.insert(*index, id.clone());
            current.set_custom_function_list(Some(make_function_list(ids, limits)?));
            add_in.set_custom_functions_with_limits(Some(current), limits)?;
        },
        CustomFunctionOperation::ReplaceId { index, id } => {
            let mut current = add_in.custom_functions().cloned().unwrap_or_default();
            let mut ids = current
                .custom_function_list()
                .map_or_else(Vec::new, |list| list.ids().to_vec());
            let Some(slot) = ids.get_mut(*index) else {
                return Err(Error::Invalid(
                    "custom-function replacement index is out of range".into(),
                ));
            };
            *slot = id.clone();
            current.set_custom_function_list(Some(make_function_list(ids, limits)?));
            add_in.set_custom_functions_with_limits(Some(current), limits)?;
        },
        CustomFunctionOperation::RemoveId { index } => {
            let mut current = add_in.custom_functions().cloned().unwrap_or_default();
            let mut ids = current
                .custom_function_list()
                .map_or_else(Vec::new, |list| list.ids().to_vec());
            if *index >= ids.len() {
                return Err(Error::Invalid(
                    "custom-function removal index is out of range".into(),
                ));
            }
            ids.remove(*index);
            current.set_custom_function_list(Some(make_function_list(ids, limits)?));
            add_in.set_custom_functions_with_limits(Some(current), limits)?;
        },
    }
    Ok(())
}

fn make_function_list(ids: Vec<String>, limits: &Limits) -> Result<CustomFunctionList> {
    if ids.len() > limits.items {
        return Err(Error::Limit {
            resource: "custom function IDs",
            max: limits.items,
            actual: ids.len(),
        });
    }
    let mut list = CustomFunctionList::new();
    for id in ids {
        list.push_id(id)?;
    }
    Ok(list)
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct ClosureRecord {
    kind: u8,
    name: String,
    before: Option<Member>,
    after: Option<Member>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Member {
    content_type: Option<String>,
    payload: Vec<u8>,
    relationships: Option<(bool, Vec<u8>)>,
}

fn encode_patch_closure(
    patch: &Patch,
    reverse: bool,
    maximum: usize,
) -> std::result::Result<Vec<u8>, PatchError> {
    let mut records = collect_patch_records(patch, false, maximum)?;
    if records.is_empty() {
        return Err(PatchError::InvalidText {
            field: "Web Extensions changed closure",
        });
    }
    let mut encoder = Encoder::new(CLOSURE_HEADER, maximum)?;
    encoder.u64(records.len())?;
    for mut record in records.drain(..) {
        if reverse {
            std::mem::swap(&mut record.before, &mut record.after);
        }
        encoder.u8(record.kind)?;
        encoder.text(&record.name)?;
        encode_member(&mut encoder, record.kind, record.before.as_ref())?;
        encode_member(&mut encoder, record.kind, record.after.as_ref())?;
    }
    Ok(encoder.finish())
}

fn collect_patch_records(
    patch: &Patch,
    include_equal: bool,
    maximum: usize,
) -> std::result::Result<Vec<ClosureRecord>, PatchError> {
    let lexical = patch.lexical.as_ref().ok_or(PatchError::InvalidText {
        field: "Web Extensions source closure",
    })?;
    preflight_patch_closure(patch, lexical, include_equal, maximum)?;
    let record_capacity = 1usize
        .checked_add(lexical.relationships.len())
        .and_then(|value| value.checked_add(patch.parts.len()))
        .ok_or(PatchError::InvalidText {
            field: "Web Extensions durable closure record count",
        })?;
    if record_capacity > MAX_CLOSURE_RECORDS {
        return Err(PatchError::InvalidText {
            field: "Web Extensions durable closure record count",
        });
    }
    let mut records = Vec::new();
    records
        .try_reserve_exact(record_capacity)
        .map_err(|_| PatchError::Allocation)?;
    let content_before = Member {
        content_type: None,
        payload: lexical.content_types.before.bytes().to_vec(),
        relationships: None,
    };
    let content_after = Member {
        content_type: None,
        payload: lexical.content_types.after.bytes().to_vec(),
        relationships: None,
    };
    if include_equal || content_before != content_after {
        records.push(ClosureRecord {
            kind: 0,
            name: "[Content_Types].xml".into(),
            before: Some(content_before),
            after: Some(content_after),
        });
    }
    for change in &lexical.relationships {
        let before = change.before.as_ref().map(|token| Member {
            content_type: None,
            payload: token.bytes().to_vec(),
            relationships: Some((token.member_present(), token.bytes().to_vec())),
        });
        let after = change.after.as_ref().map(|token| Member {
            content_type: None,
            payload: token.bytes().to_vec(),
            relationships: Some((token.member_present(), token.bytes().to_vec())),
        });
        if include_equal || before != after {
            records.push(ClosureRecord {
                kind: 1,
                name: change.owner.as_str().to_owned(),
                before,
                after,
            });
        }
    }
    for change in &patch.parts {
        if !include_equal && change.before == change.after {
            continue;
        }
        let before = part_member(change, false, lexical)?;
        let after = part_member(change, true, lexical)?;
        records.push(ClosureRecord {
            kind: 2,
            name: change.name.as_str().to_owned(),
            before,
            after,
        });
    }
    records.sort_by(|left, right| {
        left.kind
            .cmp(&right.kind)
            .then_with(|| left.name.as_bytes().cmp(right.name.as_bytes()))
    });
    Ok(records)
}

fn preflight_patch_closure(
    patch: &Patch,
    lexical: &super::LexicalChange,
    include_equal: bool,
    maximum: usize,
) -> std::result::Result<(), PatchError> {
    let mut total =
        CLOSURE_HEADER
            .len()
            .checked_add(size_of::<u64>())
            .ok_or(PatchError::InvalidText {
                field: "Web Extensions durable closure byte length",
            })?;
    if include_equal || lexical.content_types.before.bytes() != lexical.content_types.after.bytes()
    {
        charge_closure_bytes(&mut total, "[Content_Types].xml".len(), maximum)?;
        charge_closure_bytes(
            &mut total,
            lexical.content_types.before.bytes().len(),
            maximum,
        )?;
        charge_closure_bytes(
            &mut total,
            lexical.content_types.after.bytes().len(),
            maximum,
        )?;
    }
    for change in &lexical.relationships {
        let before = change.before.as_ref();
        let after = change.after.as_ref();
        if !include_equal
            && before.map(OwnedRelationships::bytes) == after.map(OwnedRelationships::bytes)
            && before.map(OwnedRelationships::member_present)
                == after.map(OwnedRelationships::member_present)
        {
            continue;
        }
        charge_closure_bytes(&mut total, change.owner.as_str().len(), maximum)?;
        charge_closure_bytes(
            &mut total,
            before.map_or(0, |token| token.bytes().len()),
            maximum,
        )?;
        charge_closure_bytes(
            &mut total,
            after.map_or(0, |token| token.bytes().len()),
            maximum,
        )?;
    }
    for change in &patch.parts {
        if !include_equal && change.before == change.after {
            continue;
        }
        charge_closure_bytes(&mut total, change.name.as_str().len(), maximum)?;
        preflight_part_member(
            &mut total,
            change.before.as_ref(),
            change,
            false,
            lexical,
            maximum,
        )?;
        preflight_part_member(
            &mut total,
            change.after.as_ref(),
            change,
            true,
            lexical,
            maximum,
        )?;
    }
    Ok(())
}

fn preflight_part_member(
    total: &mut usize,
    state: Option<&PartState>,
    change: &PartChange,
    after: bool,
    lexical: &super::LexicalChange,
    maximum: usize,
) -> std::result::Result<(), PatchError> {
    let Some(state) = state else { return Ok(()) };
    charge_closure_bytes(total, state.content_type.len(), maximum)?;
    charge_closure_bytes(total, state.data.len(), maximum)?;
    let relationship_bytes = lexical
        .relationships
        .iter()
        .find(|relationship| relationship.owner == change.name)
        .and_then(|relationship| {
            let token = if after {
                relationship.after.as_ref()
            } else {
                relationship.before.as_ref()
            }?;
            Some(token.bytes().len())
        })
        .unwrap_or_else(|| {
            state
                .relationships
                .iter()
                .map(|relationship| {
                    relationship
                        .id
                        .len()
                        .saturating_add(relationship.relationship_type.len())
                        .saturating_add(relationship.target.len())
                        .saturating_add(64)
                })
                .fold(128usize, usize::saturating_add)
        });
    charge_closure_bytes(total, relationship_bytes, maximum)?;
    Ok(())
}

fn charge_closure_bytes(
    total: &mut usize,
    amount: usize,
    maximum: usize,
) -> std::result::Result<(), PatchError> {
    *total = total.checked_add(amount).ok_or(PatchError::InvalidText {
        field: "Web Extensions durable closure byte length",
    })?;
    if *total > maximum {
        return Err(PatchError::InvalidText {
            field: "Web Extensions durable closure byte limit",
        });
    }
    Ok(())
}

fn part_member(
    change: &PartChange,
    after: bool,
    lexical: &super::LexicalChange,
) -> std::result::Result<Option<Member>, PatchError> {
    let state = if after {
        change.after.as_ref()
    } else {
        change.before.as_ref()
    };
    let Some(state) = state else { return Ok(None) };
    let relationships = lexical
        .relationships
        .iter()
        .find(|relationship| relationship.owner == change.name)
        .and_then(|relationship| {
            let token = if after {
                relationship.after.as_ref()
            } else {
                relationship.before.as_ref()
            }?;
            Some((token.member_present(), token.bytes().to_vec()))
        })
        .or_else(|| {
            Some((
                !state.relationships.is_empty(),
                canonical_relationships(state),
            ))
        });
    Ok(Some(Member {
        content_type: Some(state.content_type.clone()),
        payload: state.data.as_slice().to_vec(),
        relationships,
    }))
}

fn canonical_relationships(state: &PartState) -> Vec<u8> {
    let mut output = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for relationship in state.relationships.iter() {
        let _ = write!(
            output,
            r#"<Relationship Id="{}" Type="{}" Target="{}"{} />"#,
            escape_xml_attr(&relationship.id),
            escape_xml_attr(&relationship.relationship_type),
            escape_xml_attr(&relationship.target),
            if relationship.external {
                r#" TargetMode="External""#
            } else {
                ""
            },
        );
    }
    output.push_str("</Relationships>");
    output.into_bytes()
}

fn escape_xml_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn encode_member(
    encoder: &mut Encoder,
    kind: u8,
    member: Option<&Member>,
) -> std::result::Result<(), PatchError> {
    let Some(member) = member else {
        return encoder.u8(0);
    };
    encoder.u8(1)?;
    if kind == 2 {
        encoder.text(member.content_type.as_deref().unwrap_or_default())?;
    }
    encoder.bytes(&member.payload)?;
    if kind == 1 {
        let (present, bytes) = member
            .relationships
            .as_ref()
            .ok_or(PatchError::InvalidText {
                field: "Web Extensions relationship closure",
            })?;
        encoder.bool(*present)?;
        encoder.bytes(bytes)?;
    }
    if kind == 2 {
        let (present, bytes) = member
            .relationships
            .as_ref()
            .ok_or(PatchError::InvalidText {
                field: "Web Extensions part relationship closure",
            })?;
        encoder.bool(*present)?;
        encoder.bytes(bytes)?;
    }
    Ok(())
}

fn decode_restore<'a>(bytes: &'a [u8], limits: &Limits) -> Result<(&'a [u8], &'a [u8])> {
    let maximum = MAX_CLOSURE_BYTES;
    if bytes.len() > maximum {
        return Err(Error::Limit {
            resource: "Web Extensions durable restore bytes",
            max: maximum,
            actual: bytes.len(),
        });
    }
    let mut decoder = Decoder::new(bytes, RESTORE_HEADER, limits.total_string_bytes)?;
    let intent = decoder.bytes_slice("Web Extensions durable intent", maximum)?;
    let closure = decoder.bytes_slice("Web Extensions durable closure", maximum)?;
    if !decoder.finished() {
        return invalid("trailing Web Extensions durable restore bytes".into());
    }
    Ok((intent, closure))
}

fn decode_closure(bytes: &[u8], limits: &Limits) -> Result<Vec<ClosureRecord>> {
    let maximum = MAX_CLOSURE_BYTES;
    if bytes.len() > maximum {
        return Err(Error::Limit {
            resource: "Web Extensions durable closure bytes",
            max: maximum,
            actual: bytes.len(),
        });
    }
    let mut decoder = Decoder::new(bytes, CLOSURE_HEADER, limits.total_string_bytes)?;
    let count = decoder.u64(
        "Web Extensions durable closure records",
        MAX_CLOSURE_RECORDS.min(limits.package_parts.saturating_mul(4).max(1)),
    )?;
    // Each record has a kind byte, a length-prefixed name, and two member
    // presence bytes before any payload.  Reject impossible counts before
    // reserving the retained record vector.
    const MIN_CLOSURE_RECORD_BYTES: usize = 1 + size_of::<u64>() + 1 + 1;
    if count == 0 || count > decoder.remaining() / MIN_CLOSURE_RECORD_BYTES {
        return invalid("invalid Web Extensions durable closure record count".into());
    }
    let retained = count
        .checked_mul(size_of::<ClosureRecord>())
        .ok_or_else(|| {
            Error::Invalid("Web Extensions durable closure allocation overflow".into())
        })?;
    if retained > maximum {
        return Err(Error::Limit {
            resource: "Web Extensions durable closure records",
            max: maximum,
            actual: retained,
        });
    }
    let mut records = Vec::new();
    records
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions durable closure records",
            source,
        })?;
    let mut seen = BTreeSet::new();
    let mut total_xml_bytes = 0usize;
    let mut total_image_bytes = 0usize;
    for _ in 0..count {
        let kind = decoder.u8()?;
        if kind > 2 {
            return invalid("unknown Web Extensions durable closure member kind".into());
        }
        let name = decoder.text(
            "Web Extensions durable closure member name",
            MAX_MEMBER_NAME,
        )?;
        if kind == 0 && name != "[Content_Types].xml" {
            return invalid("invalid Web Extensions content-types closure member".into());
        }
        if kind == 1 && name != "/" {
            PackURI::new(name.clone()).map_err(Error::Uri)?;
        }
        if kind == 2 {
            if name == "/" {
                return invalid(
                    "Web Extensions durable closure part cannot be the package root".into(),
                );
            }
            PackURI::new(name.clone()).map_err(Error::Uri)?;
        }
        let before = decode_member(
            &mut decoder,
            kind,
            limits,
            &mut total_xml_bytes,
            &mut total_image_bytes,
        )?;
        let after = decode_member(
            &mut decoder,
            kind,
            limits,
            &mut total_xml_bytes,
            &mut total_image_bytes,
        )?;
        if before.is_none() && after.is_none() || before == after {
            return invalid("empty or equal Web Extensions closure record".into());
        }
        if kind == 2
            && let (Some(before), Some(after)) = (&before, &after)
            && before.content_type != after.content_type
        {
            return invalid("existing Web Extensions part content type changed".into());
        }
        let folded_name = ascii_folded_name(&name)?;
        if !seen.insert((kind, folded_name)) {
            return invalid("duplicate Web Extensions durable closure member".into());
        }
        records.push(ClosureRecord {
            kind,
            name,
            before,
            after,
        });
    }
    if !decoder.finished() {
        return invalid("trailing Web Extensions durable closure bytes".into());
    }
    Ok(records)
}

fn decode_member(
    decoder: &mut Decoder<'_>,
    kind: u8,
    limits: &Limits,
    total_xml_bytes: &mut usize,
    total_image_bytes: &mut usize,
) -> Result<Option<Member>> {
    if !decoder.bool("Web Extensions durable member presence")? {
        return Ok(None);
    }
    let content_type = if kind == 2 {
        Some(decoder.text("Web Extensions durable content type", MAX_CONTENT_TYPE)?)
    } else {
        None
    };
    let payload_is_xml =
        kind <= 1 || (kind == 2 && content_type.as_deref().is_some_and(is_xml_content_type));
    let payload_maximum = if payload_is_xml {
        limits.xml_bytes
    } else {
        limits.image_bytes
    }
    .min(MAX_CLOSURE_BYTES);
    let payload = decoder.bytes_with_total(
        "Web Extensions durable closure payload",
        payload_maximum,
        if payload_is_xml {
            total_xml_bytes
        } else {
            total_image_bytes
        },
        if payload_is_xml {
            limits.total_xml_bytes
        } else {
            limits.total_image_bytes
        },
        if payload_is_xml {
            "Web Extensions durable XML payload bytes"
        } else {
            "Web Extensions durable image payload bytes"
        },
    )?;
    let relationships = if kind == 1 {
        let present = decoder.bool("Web Extensions durable relationship presence")?;
        let bytes = decoder.bytes_with_total(
            "Web Extensions durable relationship XML",
            MAX_RELATIONSHIP_BYTES.min(limits.xml_bytes),
            total_xml_bytes,
            limits.total_xml_bytes,
            "Web Extensions durable XML payload bytes",
        )?;
        if payload != bytes {
            return invalid(
                "Web Extensions durable relationship payload disagrees with member bytes".into(),
            );
        }
        Some((present, bytes))
    } else if kind == 2 {
        let present = decoder.bool("Web Extensions durable relationship presence")?;
        let bytes = decoder.bytes_with_total(
            "Web Extensions durable relationship XML",
            MAX_RELATIONSHIP_BYTES.min(limits.xml_bytes),
            total_xml_bytes,
            limits.total_xml_bytes,
            "Web Extensions durable XML payload bytes",
        )?;
        Some((present, bytes))
    } else {
        None
    };
    Ok(Some(Member {
        content_type,
        payload,
        relationships,
    }))
}

fn is_xml_content_type(content_type: &str) -> bool {
    let bytes = content_type.as_bytes();
    content_type.eq_ignore_ascii_case("application/xml")
        || (bytes.len() >= 4 && bytes[bytes.len() - 4..].eq_ignore_ascii_case(b"+xml"))
}

fn ascii_folded_name(name: &str) -> Result<Vec<u8>> {
    let mut folded = Vec::new();
    folded
        .try_reserve_exact(name.len())
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions durable closure member names",
            source,
        })?;
    folded.extend(name.bytes().map(|byte| byte.to_ascii_lowercase()));
    Ok(folded)
}

fn validate_closure_source(
    package: &OpcPackage,
    records: &[ClosureRecord],
    _limits: &Limits,
) -> Result<()> {
    for record in records {
        let Some(expected) = record.before.as_ref() else {
            match record.kind {
                0 => return invalid("content-types source cannot be absent".into()),
                1 => {
                    let owner = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
                    match package.source_relationships(&owner) {
                        Ok(token) => {
                            if token.member_present() || !token.bytes().is_empty() {
                                return invalid(
                                    "unexpected Web Extensions relationship source".into(),
                                );
                            }
                        },
                        Err(litchi_opc::OpcError::PartNotFound(_)) => {
                            // The owner itself may be absent when the compact
                            // closure records an absent relationship member.
                        },
                        Err(error) => return Err(error.into()),
                    }
                },
                2 => {
                    let name = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
                    // Presence is a question about names. `get_part` decodes,
                    // so a present part whose payload fails to decode would
                    // read as absent and later be silently replaced.
                    if package.part_metadata(&name).is_some() {
                        return invalid("unexpected Web Extensions source part".into());
                    }
                },
                _ => return invalid("unknown Web Extensions closure kind".into()),
            }
            continue;
        };
        match record.kind {
            0 => {
                if package.source_content_types()?.bytes() != expected.payload.as_slice() {
                    return invalid("stale Web Extensions content-types source".into());
                }
            },
            1 => {
                let owner = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
                let (present, bytes) = member_relationship(expected)?;
                match package.source_relationships(&owner) {
                    Ok(token) => {
                        if token.member_present() != present || token.bytes() != bytes {
                            return invalid("stale Web Extensions relationship source".into());
                        }
                    },
                    Err(litchi_opc::OpcError::PartNotFound(_)) if !present => {
                        // A compact closure may retain the exact absent-vs-empty
                        // relationship state for a part that is itself absent.
                    },
                    Err(error) => return Err(error.into()),
                }
            },
            2 => {
                let name = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
                if package.part_metadata(&name).is_none() {
                    return Err(Error::Missing(format!(
                        "Web Extensions source part '{}'",
                        name.as_str()
                    )));
                }
                // The part is present: a decode failure is its own refusal,
                // not absence.
                let part = package.get_part(&name)?;
                if expected.content_type.as_deref() != Some(part.content_type())
                    || expected.payload.as_slice() != part.blob()
                {
                    return invalid("stale Web Extensions source part".into());
                }
                let token = package.source_relationships(&name)?;
                let (present, bytes) = member_relationship(expected)?;
                if token.member_present() != present || token.bytes() != bytes {
                    return invalid("stale Web Extensions source part relationships".into());
                }
            },
            _ => return invalid("unknown Web Extensions closure kind".into()),
        }
    }
    Ok(())
}

fn member_relationship(member: &Member) -> Result<(bool, &[u8])> {
    member
        .relationships
        .as_ref()
        .map(|(present, bytes)| (*present, bytes.as_slice()))
        .ok_or_else(|| Error::Invalid("missing Web Extensions relationship closure".into()))
}

fn reconstruct_source(
    package: &mut OpcPackage,
    records: &[ClosureRecord],
    limits: &Limits,
) -> Result<()> {
    let proof = exact_proof_package(package, records, limits)?;
    // Add new targets before installing their owner relationship tokens.
    for record in records.iter().filter(|record| record.kind == 2) {
        if record.before.is_some() || record.after.is_none() {
            continue;
        }
        let after = record
            .after
            .as_ref()
            .ok_or_else(|| Error::Invalid("missing Web Extensions target part".into()))?;
        let name = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
        let content_type = after
            .content_type
            .clone()
            .ok_or_else(|| Error::Invalid("new Web Extensions part has no content type".into()))?;
        package.add_part(Box::new(BlobPart::new(
            name,
            content_type,
            after.payload.clone(),
        )));
    }

    // Restore existing XML through a proof token.  Raw `set_blob` is reserved
    // for binary resources, where source XML provenance cannot be lost.
    for record in records.iter().filter(|record| record.kind == 2) {
        let (Some(before), Some(after)) = (&record.before, &record.after) else {
            continue;
        };
        if before.payload == after.payload {
            continue;
        }
        let name = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
        let part = package.get_part(&name).map_err(|error| {
            Error::Missing(format!(
                "Web Extensions source reconstruction part '{}' is unavailable: {error}",
                name.as_str()
            ))
        })?;
        if is_xml_content(part.content_type(), &name) {
            let replacement = proof.source_xml_part(&name).map_err(|error| {
                Error::Missing(format!(
                    "Web Extensions proof XML part '{}' is unavailable: {error}",
                    name.as_str()
                ))
            })?;
            package.try_replace_owned_xml_part(before.payload.as_slice(), replacement)?;
        } else {
            package.get_part_mut(&name)?.set_blob(after.payload.clone());
        }
    }

    // Relationship bytes are installed only through tokens parsed by OPC.
    for record in records.iter().filter(|record| record.kind == 1) {
        let Some(after) = record.after.as_ref() else {
            continue;
        };
        let owner = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
        let current = package.source_relationships(&owner).map_err(|error| {
            Error::Missing(format!(
                "Web Extensions relationship owner '{}' is unavailable during source reconstruction: {error}",
                owner.as_str()
            ))
        })?;
        let replacement = proof.source_relationships(&owner)?;
        let (present, bytes) = member_relationship(after)?;
        if replacement.member_present() != present || replacement.bytes() != bytes {
            return invalid("Web Extensions relationship proof token mismatch".into());
        }
        if current != replacement {
            package.try_replace_relationships(&current, &replacement)?;
        }
    }
    // Part relationship tokens may be changed without a separate owner record
    // in a compact closure, so install every target token carried by a part.
    for record in records.iter().filter(|record| record.kind == 2) {
        let Some(after) = record.after.as_ref() else {
            continue;
        };
        let name = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
        if !package_has_part(package, &record.name)? {
            continue;
        }
        let current = package.source_relationships(&name).map_err(|error| {
            Error::Missing(format!(
                "Web Extensions part relationship owner '{}' is unavailable during source reconstruction: {error}",
                name.as_str()
            ))
        })?;
        let replacement = proof.source_relationships(&name).map_err(|error| {
            Error::Missing(format!(
                "Web Extensions proof part relationship owner '{}' is unavailable: {error}",
                name.as_str()
            ))
        })?;
        let (present, bytes) = member_relationship(after)?;
        if replacement.member_present() != present || replacement.bytes() != bytes {
            return invalid("Web Extensions part relationship proof token mismatch".into());
        }
        if current != replacement {
            package.try_replace_relationships(&current, &replacement)?;
        }
    }

    // Removed parts are safe only after all target owners have been installed.
    for record in records.iter().filter(|record| record.kind == 2) {
        if record.after.is_some() {
            continue;
        }
        let name = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
        if has_incoming(package, &name)? || !package.remove_part(&name) {
            return invalid("Web Extensions source reconstruction leaves an incoming edge".into());
        }
    }

    if let Some(content) = records.iter().find(|record| record.kind == 0) {
        let after = content
            .after
            .as_ref()
            .ok_or_else(|| Error::Invalid("missing Web Extensions target content types".into()))?;
        let replacement = proof.source_content_types()?;
        if replacement.bytes() != after.payload.as_slice() {
            return invalid("Web Extensions content-types proof token mismatch".into());
        }
        let current = package.source_content_types()?;
        if current.bytes() != replacement.bytes() {
            package.try_replace_content_types(current.bytes(), &replacement)?;
        }
    }
    Ok(())
}

fn package_has_part(package: &OpcPackage, name: &str) -> Result<bool> {
    let name = PackURI::new(name.to_owned()).map_err(Error::Uri)?;
    // Without decoding: a present part whose payload fails to decode is still
    // present.
    Ok(package.part_metadata(&name).is_some())
}

fn is_xml_content(content_type: &str, name: &PackURI) -> bool {
    content_type.eq_ignore_ascii_case("application/xml")
        || content_type.ends_with("+xml")
        || name.as_str().rsplit('/').next().is_some_and(|part| {
            part.len() >= 4 && part[part.len() - 4..].eq_ignore_ascii_case(".xml")
        })
}

fn has_incoming(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    for relationship in package.rels().iter() {
        if !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|name| name.is_equivalent_to(target))
        {
            return Ok(true);
        }
    }
    for owner in package.iter_parts() {
        for relationship in owner.rels().iter() {
            if !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name.is_equivalent_to(target))
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn exact_proof_package(
    package: &OpcPackage,
    records: &[ClosureRecord],
    limits: &Limits,
) -> Result<OpcPackage> {
    let maximum = limits
        .total_xml_bytes
        .max(limits.total_image_bytes)
        .min(MAX_CLOSURE_BYTES);
    if maximum == 0 {
        return invalid("Web Extensions durable proof limit is zero".into());
    }
    let content_types_owned = records
        .iter()
        .find(|record| record.kind == 0)
        .and_then(|record| record.after.as_ref())
        .map(|member| member.payload.clone())
        .unwrap_or(package.source_content_types()?.bytes().to_vec());
    if content_types_owned.is_empty() {
        return invalid("Web Extensions proof has no content-types XML".into());
    }
    let content_types = content_types_owned.as_slice();
    let mut names = BTreeSet::new();
    for part in package.iter_parts() {
        names.insert(part.partname().as_str().to_owned());
    }
    for record in records.iter().filter(|record| record.kind == 2) {
        if record.after.is_some() {
            names.insert(record.name.clone());
        } else {
            names.remove(&record.name);
        }
    }
    if names.len() > limits.package_parts {
        return Err(Error::Limit {
            resource: "Web Extensions durable proof parts",
            max: limits.package_parts,
            actual: names.len(),
        });
    }
    let root_name = PackURI::new("/").map_err(Error::Uri)?;
    let root_current = package.source_relationships(&root_name)?;
    let root_target = relation_target(records, "/", Some(&root_current))?;
    let mut total = content_types.len().saturating_add(PROOF_ENTRY_OVERHEAD);
    for name in &names {
        let part = proof_part(package, records, name)?;
        total = total
            .checked_add(part.content_type.len())
            .and_then(|value| value.checked_add(part.payload.len()))
            .and_then(|value| value.checked_add(part.relationships.1.len()))
            .and_then(|value| value.checked_add(name.len()))
            .and_then(|value| value.checked_add(PROOF_ENTRY_OVERHEAD))
            .ok_or(Error::Limit {
                resource: "Web Extensions durable proof bytes",
                max: maximum,
                actual: usize::MAX,
            })?;
        if total > maximum {
            return Err(Error::Limit {
                resource: "Web Extensions durable proof bytes",
                max: maximum,
                actual: total,
            });
        }
    }
    let mut writer = PhysPkgWriter::new();
    let content_name = PackURI::new("/[Content_Types].xml").map_err(Error::Uri)?;
    writer.write_stored(&content_name, content_types)?;
    if root_target.0 {
        let root_rels = root_name.rels_uri().map_err(Error::Uri)?;
        writer.write_stored(&root_rels, &root_target.1)?;
    }
    for name in &names {
        let proof = proof_part(package, records, name)?;
        let part_name = PackURI::new(name.clone()).map_err(Error::Uri)?;
        writer.write_stored(&part_name, &proof.payload)?;
    }
    for name in &names {
        let proof = proof_part(package, records, name)?;
        if proof.relationships.0 {
            let part_name = PackURI::new(name.clone()).map_err(Error::Uri)?;
            let rels = part_name.rels_uri().map_err(Error::Uri)?;
            writer.write_stored(&rels, &proof.relationships.1)?;
        }
    }
    Ok(OpcPackage::from_bytes(&writer.finish()?)?)
}

struct ProofPart<'a> {
    content_type: String,
    payload: Vec<u8>,
    relationships: (bool, Vec<u8>),
    _marker: std::marker::PhantomData<&'a ()>,
}

fn proof_part(
    package: &OpcPackage,
    records: &[ClosureRecord],
    name: &str,
) -> Result<ProofPart<'static>> {
    let changed = records
        .iter()
        .find(|record| record.kind == 2 && record.name == name);
    if let Some(record) = changed {
        let after = record
            .after
            .as_ref()
            .ok_or_else(|| Error::Invalid("missing Web Extensions proof target part".into()))?;
        let content_type = after
            .content_type
            .clone()
            .ok_or_else(|| Error::Invalid("missing Web Extensions proof content type".into()))?;
        let relationships = member_relationship(after)?.to_owned();
        return Ok(ProofPart {
            content_type,
            payload: after.payload.clone(),
            relationships: (relationships.0, relationships.1.to_vec()),
            _marker: std::marker::PhantomData,
        });
    }
    let part_name = PackURI::new(name.to_owned()).map_err(Error::Uri)?;
    let part = package.get_part(&part_name).map_err(|error| {
        Error::Missing(format!(
            "Web Extensions proof source fallback part '{}' is unavailable: {error}",
            part_name.as_str()
        ))
    })?;
    let relationships = package.source_relationships(&part_name)?;
    Ok(ProofPart {
        content_type: part.content_type().to_owned(),
        payload: Vec::new(),
        relationships: (
            relationships.member_present(),
            relationships.bytes().to_vec(),
        ),
        _marker: std::marker::PhantomData,
    })
}

fn relation_target(
    records: &[ClosureRecord],
    owner: &str,
    current: Option<&OwnedRelationships>,
) -> Result<(bool, Vec<u8>)> {
    if let Some(record) = records
        .iter()
        .find(|record| record.kind == 1 && record.name == owner)
    {
        let after = record
            .after
            .as_ref()
            .ok_or_else(|| Error::Invalid("missing Web Extensions relationship target".into()))?;
        let (present, bytes) = member_relationship(after)?;
        return Ok((present, bytes.to_vec()));
    }
    let current =
        current.ok_or_else(|| Error::Invalid("missing Web Extensions relationships".into()))?;
    Ok((current.member_present(), current.bytes().to_vec()))
}
