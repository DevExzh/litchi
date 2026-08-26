//! Native Pop-Up Menu graph transitions.
//!
//! This module is deliberately below the public Pop-Up Menu facade.  It takes
//! already-resolved native routes and returns private archive bytes; selector,
//! package metadata, ZIP, preview, and publication policy remain owned by the
//! facade.  The compatibility entry point accepts one decompressed IWA
//! archive; the split entry point accepts a complete set of current members
//! and projects its private merged candidate back to those members.

use std::fmt;

use litchi_iwa_common::{decode_varint_from_bytes, wire::WireView};
use litchi_iwa_core::archive::{
    ArchiveReferenceKind, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceVisitor, FieldObjectReferenceTransition, ObjectReferenceTransition,
};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldType, Limits, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{
    numbers_table_cell_control_codec as control_codec,
    numbers_table_cell_pop_up_menu_codec as popup_codec,
    numbers_table_cell_storage_codec as storage_codec,
};
use litchi_numbers_wire::{
    BncCell,
    popup_menu::{
        PopUpMenuBncState, PopUpMenuRewriteError, PopUpMenuRewriteOptions, prepare_pop_up_menu_bnc,
    },
};

use super::table_cell_pop_up_menu::{Path, TransactionBudget};

const TABLE_MODEL_TYPE: u32 = 6_001;
const TILE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_TYPE: u32 = 6_005;
const POPUP_MODEL_TYPE: u32 = 6_206;
const LIST_STRING: i32 = 1;
const LIST_FORMAT: i32 = 2;
const LIST_CONTROL_CELL_SPEC: i32 = 12;

/// A native Pop-Up Menu value after semantic validation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct NativePopUpValue<'items> {
    /// Menu labels, excluding the required native NIL sentinel.
    pub(super) items: &'items [&'items str],
    /// Native CellSpecArchive chooser flag.
    pub(super) starts_with_first: bool,
    /// Existing String table key for the selected first item, if any.
    pub(super) first_item_string_identifier: Option<u32>,
}

/// A fully resolved native graph for the compatibility single-member path.
///
/// Every identifier must occur in the supplied archive exactly once.  The
/// package owner resolves these identifiers from its rooted selector before
/// calling this module; retaining identifiers here keeps the native helper
/// independent from facade selectors and generated model types.
#[derive(Debug, Clone, Copy)]
pub(super) struct NativePopUpInput<'source> {
    pub(super) archive: &'source Archive,
    /// Index of the physical component containing `archive`.  Keeping this
    /// beside the member name makes the private output usable by owners that
    /// stage more than one sidecar in one transaction; a name alone is not a
    /// safe component identity when effective locators are involved.
    pub(super) component_index: usize,
    pub(super) model_identifier: u64,
    pub(super) tile_identifier: u64,
    pub(super) tile_row: u32,
    pub(super) tile_column: u32,
    pub(super) control_table_identifier: u64,
    pub(super) format_table_identifier: u64,
    /// The physical component/member name containing `archive`.
    pub(super) member_name: &'source str,
    /// FormatStructArchive bytes used when a format-list entry is created.
    pub(super) format_payload: &'source [u8],
    /// Identifier reserved by the package metadata allocator for a new
    /// PopUpMenuModel object. It is unused when an identical model is reused.
    pub(super) new_popup_model_identifier: Option<u64>,
    pub(super) desired: Option<NativePopUpValue<'source>>,
    pub(super) limits: Limits,
    pub(super) path: Path,
}

/// A physical IWA member supplied to the multi-member native transition.
///
/// The archive is borrowed from the package source and is never mutated in
/// place.  A member is identified by both its component index and effective
/// member name; either value alone is insufficient when a package contains
/// versioned or explicitly-located component records.
#[derive(Debug, Clone, Copy)]
pub(super) struct NativePopUpMember<'source> {
    pub(super) archive: &'source Archive,
    pub(super) component_index: usize,
    pub(super) member_name: &'source str,
}

/// A rooted object route into one of [`NativePopUpMember`]s.
#[derive(Debug, Clone, Copy)]
pub(super) struct NativePopUpObjectRoute {
    pub(super) member_index: usize,
    pub(super) identifier: u64,
}

/// The resolved physical graph for a split Pop-Up Menu transaction.
///
/// The model, tile, and list objects may live in different current component
/// members.  The native transition merges only their parsed object graphs in
/// a private working archive, performs the same strict lifecycle/census logic
/// as the single-member path, then splits the candidate back along the source
/// ownership map.  This keeps cross-member references visible to refcount and
/// cull checks without ever publishing a merged archive.
#[derive(Debug, Clone, Copy)]
pub(super) struct NativePopUpGraphInput<'source> {
    /// All current physical members that can contain rooted table/storage
    /// objects, not merely the members named by the selected cell.  The
    /// merged working archive uses this complete set to census BNC/list
    /// references before a refcount decrement or popup cull.
    pub(super) members: &'source [NativePopUpMember<'source>],
    pub(super) model: NativePopUpObjectRoute,
    pub(super) tile: NativePopUpObjectRoute,
    pub(super) control_table_identifier: u64,
    pub(super) format_table_identifier: u64,
    /// Member in which a newly-created PopUpMenuModel is allowed to live.
    /// Existing models are assigned to their resolved source member.
    pub(super) creation_member_index: usize,
    pub(super) tile_row: u32,
    pub(super) tile_column: u32,
    pub(super) format_payload: &'source [u8],
    pub(super) new_popup_model_identifier: Option<u64>,
    pub(super) desired: Option<NativePopUpValue<'source>>,
    pub(super) limits: Limits,
    pub(super) path: Path,
}

/// One private physical-member candidate produced by a native owner.
///
/// Component index and effective member name are both retained deliberately:
/// a ZIP member name is only a locator, while the component index is the
/// identity used for metadata save-token selection.  The two values must
/// agree with the source catalog before publication.
#[derive(Debug, PartialEq, Eq)]
pub(super) struct NativeMemberEdit {
    pub(super) component_index: usize,
    pub(super) member_name: String,
    pub(super) member_bytes: Vec<u8>,
}

impl NativeMemberEdit {
    pub(super) fn new(
        component_index: usize,
        member_name: impl Into<String>,
        member_bytes: Vec<u8>,
    ) -> Self {
        Self {
            component_index,
            member_name: member_name.into(),
            member_bytes,
        }
    }
}

/// Deduplicated private native edits for one atomic package transaction.
///
/// A component/member may occur at most once.  Identical duplicate edits are
/// coalesced, but conflicting bytes are rejected rather than letting the ZIP
/// reassembler's last-write-wins behavior hide an ownership bug.  The same
/// rule applies to a member name that was resolved to two component indices.
#[derive(Debug, PartialEq, Eq)]
pub(super) struct NativeControlOutput {
    pub(super) edits: Vec<NativeMemberEdit>,
}

impl NativeControlOutput {
    pub(super) fn from_edits(edits: Vec<NativeMemberEdit>) -> Result<Self> {
        let mut deduplicated: Vec<NativeMemberEdit> = Vec::new();
        deduplicated
            .try_reserve_exact(edits.len())
            .map_err(|_| NativePopUpError::Allocation)?;
        for edit in edits {
            if edit.member_name.is_empty() {
                return Err(NativePopUpError::InvalidSource);
            }
            if let Some(existing) = deduplicated
                .iter()
                .find(|existing| existing.member_name == edit.member_name)
            {
                if existing.component_index != edit.component_index
                    || existing.member_bytes != edit.member_bytes
                {
                    return Err(NativePopUpError::UnsupportedDependency);
                }
                continue;
            }
            deduplicated.push(edit);
        }
        deduplicated.sort_unstable_by(|left, right| {
            left.component_index
                .cmp(&right.component_index)
                .then(left.member_name.cmp(&right.member_name))
        });
        if deduplicated.is_empty() {
            return Err(NativePopUpError::InvalidSource);
        }
        Ok(Self {
            edits: deduplicated,
        })
    }

    pub(super) fn single(
        component_index: usize,
        member_name: impl Into<String>,
        member_bytes: Vec<u8>,
    ) -> Result<Self> {
        Self::from_edits(vec![NativeMemberEdit::new(
            component_index,
            member_name,
            member_bytes,
        )])
    }

    pub(super) fn member_names(&self) -> impl Iterator<Item = &str> + '_ {
        self.edits.iter().map(|edit| edit.member_name.as_str())
    }
}

/// Owned native output returned to the package transaction.
///
/// Each `member_edits` payload is a decompressed IWA candidate. The caller
/// compresses each member independently and places the resulting edits in the
/// final private ZIP candidate. `cell_bytes` is exposed separately so callers
/// can perform an object-level locality assertion for the selected BNC cell.
#[derive(Debug, PartialEq, Eq)]
pub(super) struct NativePopUpOutput {
    /// Deduplicated physical edits.  The compatibility path normally emits
    /// one edit; the split path emits one edit for each changed member so a
    /// format/control graph can stage all changed members atomically.
    pub(super) member_edits: NativeControlOutput,
    pub(super) cell_bytes: Vec<u8>,
    pub(super) format_identifier: Option<u32>,
    pub(super) control_cell_spec_identifier: Option<u32>,
    pub(super) popup_model_identifier: Option<u64>,
    pub(super) added_object_identifiers: Vec<u64>,
    pub(super) removed_object_identifiers: Vec<u64>,
    pub(super) changed_member_names: Vec<String>,
}

/// Candidate archive transition before it is serialized into one or more
/// physical members.  Keeping the parsed candidate here is what allows the
/// split-member wrapper to preserve each source member's object ownership and
/// object order.
#[derive(Debug)]
struct NativePopUpTransition {
    candidate: Archive,
    cell_bytes: Vec<u8>,
    format_identifier: Option<u32>,
    control_cell_spec_identifier: Option<u32>,
    popup_model_identifier: Option<u64>,
    added_object_identifiers: Vec<u64>,
    removed_object_identifiers: Vec<u64>,
}

/// Native graph failure. The facade maps this to its content-redacted error
/// enum; no native payload is included in the error value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum NativePopUpError {
    InvalidSource,
    UnsupportedDependency,
    Allocation,
    Limit,
    Codec,
    Archive,
}

impl fmt::Display for NativePopUpError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InvalidSource => "invalid Pop-Up Menu native source",
            Self::UnsupportedDependency => "unsupported Pop-Up Menu native dependency",
            Self::Allocation => "Pop-Up Menu native allocation failed",
            Self::Limit => "Pop-Up Menu native limit exceeded",
            Self::Codec => "Pop-Up Menu native codec rejected the source",
            Self::Archive => "Pop-Up Menu native archive rewrite failed",
        })
    }
}

impl std::error::Error for NativePopUpError {}

pub(super) type Result<T> = std::result::Result<T, NativePopUpError>;

/// Plan and execute one compatibility single-member Pop-Up Menu transition.
///
/// The function validates the model/DataStore field-21 route, exactly one
/// control and format list, the selected tile/BNC cell, every existing popup
/// CellSpec, and every type-6206 model it encounters. It then uses the strict
/// prepared codecs for BNC, CellSpec, PopUpMenuModel, and TableDataList. Raw
/// protobuf fields outside the selected list entry/row are copied verbatim.
///
/// Existing menu/spec objects are reused when their canonical semantic bytes
/// match. A new model is created only when no identical model exists and the
/// caller supplied a collision-free metadata-reserved identifier. A model is
/// culled only after a complete same-archive inbound scan proves that no
/// remaining control-list entry references it. Empty control lists are kept.
fn rewrite_native_popup_menu_transition(
    input: NativePopUpInput<'_>,
    budget: &mut TransactionBudget,
    preserve_aggregate_only: bool,
) -> Result<NativePopUpTransition> {
    let source = input.archive;
    let decoded_bytes = source
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .map(|message| message.data.len())
        .sum::<usize>();
    budget
        .charge_archive(source, decoded_bytes, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    validate_unique_route(source, input.model_identifier)?;
    validate_unique_route(source, input.tile_identifier)?;
    validate_unique_route(source, input.control_table_identifier)?;
    validate_unique_route(source, input.format_table_identifier)?;
    if input.model_identifier == input.tile_identifier
        || input.model_identifier == input.control_table_identifier
        || input.model_identifier == input.format_table_identifier
        || input.tile_identifier == input.control_table_identifier
        || input.tile_identifier == input.format_table_identifier
    {
        return Err(NativePopUpError::UnsupportedDependency);
    }

    let model_object = unique_object(source, input.model_identifier)?;
    let model_message_index = unique_message_index(model_object, TABLE_MODEL_TYPE)?;
    let model_payload = &model_object.messages[model_message_index].data;
    let model_options = budget.residual_storage_options(model_payload);
    let (model, model_report) =
        storage_codec::decode_table_model_with_report(model_payload, model_options)
            .map_err(|_| NativePopUpError::Codec)?;
    budget
        .charge_storage_decode_report(model_report, true, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    let (store, store_report) = storage_codec::decode_data_store_with_report(
        model.base_data_store(),
        budget.residual_storage_options(model.base_data_store()),
    )
    .map_err(|_| NativePopUpError::Codec)?;
    budget
        .charge_storage_decode_report(store_report, true, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    if store
        .control_cell_spec_table()
        .map(|reference| reference.identifier())
        != Some(input.control_table_identifier)
        || store.format_table().map(|reference| reference.identifier())
            != Some(input.format_table_identifier)
    {
        return Err(NativePopUpError::InvalidSource);
    }

    let tile_identifier = tile_identifier(
        store.tiles(),
        input.tile_identifier,
        input.tile_row,
        budget.residual_storage_options(store.tiles()),
        budget,
        input.path,
    )?;
    if tile_identifier != input.tile_identifier {
        return Err(NativePopUpError::InvalidSource);
    }
    let tile_object = unique_object(source, input.tile_identifier)?;
    let tile_message_index = unique_message_index(tile_object, TILE_TYPE)?;
    let tile_payload = &tile_object.messages[tile_message_index].data;
    let cell_source = tile_cell(tile_payload, input.tile_row, input.tile_column)?;
    let cell = BncCell::parse(cell_source).map_err(|_| NativePopUpError::InvalidSource)?;
    let old_format = cell.format_identifier();
    let old_control = cell.control_cell_spec_identifier();
    if old_format.is_some() != old_control.is_some() {
        return Err(NativePopUpError::InvalidSource);
    }

    let control_list_object = unique_object(source, input.control_table_identifier)?;
    let control = unique_list_message(
        control_list_object,
        LIST_CONTROL_CELL_SPEC,
        model_options,
        budget,
        input.path,
    )?;
    let format_list_object = unique_object(source, input.format_table_identifier)?;
    let format = unique_list_message(
        format_list_object,
        LIST_FORMAT,
        model_options,
        budget,
        input.path,
    )?;
    let bnc_references = census_bnc_references(source, budget, input.path)?;
    validate_bnc_refcounts(&bnc_references, format.payload, LIST_FORMAT, model_options)?;
    validate_bnc_refcounts(
        &bnc_references,
        control.payload,
        LIST_CONTROL_CELL_SPEC,
        model_options,
    )?;
    if let Some(format_key) = old_format {
        let format_entries = decode_list_entries(format.payload, model_options)?;
        let entry = format_entries
            .entries
            .iter()
            .find(|entry| entry.key == format_key)
            .ok_or(NativePopUpError::InvalidSource)?;
        if entry.ref_count == 0 || entry.payload_kind != PayloadKind::Format {
            return Err(NativePopUpError::InvalidSource);
        }
    }
    // A string list is part of the rooted DataStore graph even when this
    // transition does not need to add a string key. Decode it to reject a
    // duplicate/malformed list owner before any candidate allocation.
    let string_object = unique_object(source, store.string_table().identifier())?;
    let _ = unique_list_message(
        string_object,
        LIST_STRING,
        model_options,
        budget,
        input.path,
    )?;

    let old_state = inspect_old_state(
        source,
        control.payload,
        old_control,
        old_format,
        model_options,
    )?;
    if input.desired.is_none() {
        validate_control_list_payloads(control.payload, model_options, budget, input.path)?;
    }
    // Only models reachable from the rooted control-list graph may be reused.
    // A semantically identical orphan 6206 object is not a safe deduplication
    // candidate: it may be versioned, unregistered, or owned by another
    // component in Metadata.iwa.
    let rooted_popup_identifiers = popup_references_from_list(control.payload)?;
    let desired_state = match input.desired {
        None => None,
        Some(desired) => Some(prepare_desired_state(
            source,
            control.payload,
            desired,
            &rooted_popup_identifiers,
            input.new_popup_model_identifier,
            model_options,
            budget,
        )?),
    };

    // All native graph vectors, BNC output, archive serialization, and
    // compression remain private candidates. Reserve a conservative envelope
    // before the first clone/output allocation; the exact codec reports are
    // charged by their prepared plans where available.
    let maximum_archive_bytes = input.limits.max_archive_bytes();
    let maximum_compressed_bytes = SnappyStream::maximum_compressed_len(maximum_archive_bytes)
        .map_err(|_| NativePopUpError::Limit)?;
    budget
        .charge_allocations(12, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    budget
        .charge_transaction_work(
            maximum_archive_bytes.saturating_add(maximum_compressed_bytes),
            input.path,
        )
        .map_err(|_| NativePopUpError::Limit)?;

    let mut candidate = source.clone();
    let mut added = Vec::new();
    let mut removed = Vec::new();

    let desired_format_payload = input.format_payload;
    let (format_bytes, format_key) = rewrite_list_for_format(
        format.payload,
        old_state.map(|state| state.format_key),
        desired_state.as_ref().map(|state| state.format_key),
        desired_state.as_ref().map(|_| desired_format_payload),
        model_options,
        budget,
        input.path,
    )?;
    let (control_bytes, control_key, desired_popup_identifier) = rewrite_control_list(
        control.payload,
        old_state,
        desired_state.as_ref(),
        model_options,
        budget,
        input.path,
    )?;

    replace_message_preserving_header(
        &mut candidate,
        input.format_table_identifier,
        format.message_index,
        format_bytes,
        input.limits,
    )?;
    replace_control_message_with_transition(
        &mut candidate,
        input.control_table_identifier,
        control.message_index,
        control.payload,
        control_bytes,
        input.limits,
        preserve_aggregate_only,
    )?;

    let desired_bnc = match desired_state.as_ref() {
        Some(state) => PopUpMenuBncState::set(
            format_key.ok_or(NativePopUpError::InvalidSource)?,
            control_key.ok_or(NativePopUpError::InvalidSource)?,
            state.first_item_string_identifier,
        ),
        None => PopUpMenuBncState::reset(),
    };
    let bnc_plan = prepare_pop_up_menu_bnc(cell_source, desired_bnc, bnc_options(cell_source))
        .map_err(map_bnc_error)?;
    let bnc_limits = bnc_plan.execution_requirements().exact_limits();
    let bnc_output = bnc_plan
        .execute(bnc_limits)
        .map_err(map_bnc_error)?
        .into_bytes();
    let tile_bytes = patch_tile_cell(tile_payload, input.tile_row, input.tile_column, &bnc_output)?;
    replace_message_preserving_header(
        &mut candidate,
        input.tile_identifier,
        tile_message_index,
        tile_bytes,
        input.limits,
    )?;

    if let Some(desired) = desired_state.as_ref() {
        if desired.new_popup_payload.is_some() {
            let identifier = desired
                .popup_identifier
                .ok_or(NativePopUpError::InvalidSource)?;
            if candidate.object(identifier).is_some() {
                return Err(NativePopUpError::InvalidSource);
            }
            let object = ArchiveObject::new_with_limits(
                identifier,
                vec![RawMessage {
                    type_: POPUP_MODEL_TYPE,
                    data: desired
                        .new_popup_payload
                        .clone()
                        .ok_or(NativePopUpError::InvalidSource)?,
                }],
                input.limits,
            )
            .map_err(|_| NativePopUpError::Archive)?;
            candidate
                .objects
                .try_reserve_exact(1)
                .map_err(|_| NativePopUpError::Allocation)?;
            candidate.objects.push(object);
            added.push(identifier);
        }
    }

    if let Some(old) = old_state.and_then(|state| state.popup_identifier) {
        if !archive_has_popup_reference(&candidate, old)? {
            candidate
                .remove_object(old)
                .ok_or(NativePopUpError::InvalidSource)?;
            removed.push(old);
        }
    }

    // Candidate archive locality: every object not explicitly authorized by
    // this transition must remain source-identical, including ArchiveInfo,
    // aggregate/FieldInfo references, unknown header framing, and raw payload
    // bytes.  A semantic reread alone would miss an unrelated registry/object
    // mutation caused by a broad replacement loop.
    let mut changed_identifiers = vec![
        input.tile_identifier,
        input.control_table_identifier,
        input.format_table_identifier,
    ];
    changed_identifiers.extend(added.iter().copied());
    changed_identifiers.extend(removed.iter().copied());
    for source_object in &source.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(NativePopUpError::InvalidSource)?;
        if removed.contains(&identifier) || changed_identifiers.contains(&identifier) {
            continue;
        }
        let candidate_object = candidate
            .object(identifier)
            .ok_or(NativePopUpError::InvalidSource)?;
        if !source_object.same_content_ignoring_offsets(candidate_object) {
            return Err(NativePopUpError::InvalidSource);
        }
    }
    for candidate_object in &candidate.objects {
        let identifier = candidate_object
            .archive_info
            .identifier
            .ok_or(NativePopUpError::InvalidSource)?;
        if source.object(identifier).is_none() && !added.contains(&identifier) {
            return Err(NativePopUpError::InvalidSource);
        }
    }

    // The list refcounts and the actual BNC references must agree after the
    // private graph transition as well.  This catches a stale refcount before
    // the candidate leaves the native owner, including the final-reset/cull
    // branch where the selected list entry disappears.
    let candidate_bnc_references = census_bnc_references(&candidate, budget, input.path)?;
    let candidate_format = unique_list_message(
        unique_object(&candidate, input.format_table_identifier)?,
        LIST_FORMAT,
        model_options,
        budget,
        input.path,
    )?;
    let candidate_control = unique_list_message(
        unique_object(&candidate, input.control_table_identifier)?,
        LIST_CONTROL_CELL_SPEC,
        model_options,
        budget,
        input.path,
    )?;
    validate_bnc_refcounts(
        &candidate_bnc_references,
        candidate_format.payload,
        LIST_FORMAT,
        model_options,
    )?;
    validate_bnc_refcounts(
        &candidate_bnc_references,
        candidate_control.payload,
        LIST_CONTROL_CELL_SPEC,
        model_options,
    )?;

    Ok(NativePopUpTransition {
        candidate,
        cell_bytes: bnc_output,
        format_identifier: format_key,
        control_cell_spec_identifier: control_key,
        popup_model_identifier: desired_popup_identifier,
        added_object_identifiers: added,
        removed_object_identifiers: removed,
    })
}

/// Plan and serialize one same-component Pop-Up Menu transition.
///
/// This is retained as the compatibility entry point for callers that have
/// already proved all graph objects share one member.  Split-component callers
/// should use [`rewrite_native_popup_menu_multi`], which emits one edit per
/// changed member.
pub(super) fn rewrite_native_popup_menu(
    input: NativePopUpInput<'_>,
    budget: &mut TransactionBudget,
) -> Result<NativePopUpOutput> {
    let transition = rewrite_native_popup_menu_transition(input, budget, false)?;
    let archive_bytes = transition
        .candidate
        .to_bytes_with_limits(input.limits)
        .map_err(|_| NativePopUpError::Archive)?;
    let member_edits =
        NativeControlOutput::single(input.component_index, input.member_name, archive_bytes)?;
    let changed_members = member_edits.member_names().map(str::to_owned).collect();
    Ok(NativePopUpOutput {
        member_edits,
        cell_bytes: transition.cell_bytes,
        format_identifier: transition.format_identifier,
        control_cell_spec_identifier: transition.control_cell_spec_identifier,
        popup_model_identifier: transition.popup_model_identifier,
        added_object_identifiers: transition.added_object_identifiers,
        removed_object_identifiers: transition.removed_object_identifiers,
        changed_member_names: changed_members,
    })
}

/// Plan and execute a split-component Pop-Up Menu transition.
///
/// The strict graph algorithm above intentionally operates on one `Archive`:
/// that makes global BNC/list/model refcount and cull checks straightforward.
/// This adapter constructs a private object-only view by concatenating the
/// resolved current members, runs the algorithm once, and then projects the
/// candidate back to the exact source member that owned every object.  Object
/// identifiers must be globally unique across the supplied members; accepting
/// an alias here would make a cull or a refcount transition ambiguous.
pub(super) fn rewrite_native_popup_menu_multi(
    input: NativePopUpGraphInput<'_>,
    budget: &mut TransactionBudget,
) -> Result<NativePopUpOutput> {
    let source_object_count = input
        .members
        .iter()
        .try_fold(0usize, |total, member| {
            total.checked_add(member.archive.objects.len())
        })
        .ok_or(NativePopUpError::InvalidSource)?;
    let source_payload_bytes = input
        .members
        .iter()
        .flat_map(|member| member.archive.objects.iter())
        .flat_map(|object| object.messages.iter())
        .try_fold(0usize, |total, message| {
            total.checked_add(message.data.len())
        })
        .ok_or(NativePopUpError::InvalidSource)?;
    // The merged working view and the per-member projection each retain an
    // object slot.  Charge those private scratch allocations before either
    // phase starts; the ordinary single-member path has no equivalent copy.
    let merge_allocation_count = source_object_count
        .checked_mul(2)
        .and_then(|count| count.checked_add(input.members.len()))
        .ok_or(NativePopUpError::InvalidSource)?;
    budget
        .charge_allocations(merge_allocation_count, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    budget
        .charge_scratch_bytes(source_payload_bytes, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    let (working, ownership) = merge_popup_members(&input)?;
    let placement = input
        .members
        .get(input.creation_member_index)
        .ok_or(NativePopUpError::InvalidSource)?;
    let model_member = input
        .members
        .get(input.model.member_index)
        .ok_or(NativePopUpError::InvalidSource)?;
    let tile_member = input
        .members
        .get(input.tile.member_index)
        .ok_or(NativePopUpError::InvalidSource)?;
    if model_member
        .archive
        .objects
        .iter()
        .filter(|object| object.archive_info.identifier == Some(input.model.identifier))
        .count()
        != 1
        || tile_member
            .archive
            .objects
            .iter()
            .filter(|object| object.archive_info.identifier == Some(input.tile.identifier))
            .count()
            != 1
    {
        return Err(NativePopUpError::InvalidSource);
    }
    if input.model.identifier == input.tile.identifier {
        return Err(NativePopUpError::UnsupportedDependency);
    }
    let synthetic = NativePopUpInput {
        archive: &working,
        component_index: placement.component_index,
        model_identifier: input.model.identifier,
        tile_identifier: input.tile.identifier,
        tile_row: input.tile_row,
        tile_column: input.tile_column,
        control_table_identifier: input.control_table_identifier,
        format_table_identifier: input.format_table_identifier,
        member_name: placement.member_name,
        format_payload: input.format_payload,
        new_popup_model_identifier: input.new_popup_model_identifier,
        desired: input.desired,
        limits: input.limits,
        path: input.path,
    };
    let transition = rewrite_native_popup_menu_transition(synthetic, budget, true)?;
    let member_edits = split_popup_candidate(
        &input,
        &ownership,
        &transition.candidate,
        &transition.added_object_identifiers,
        &transition.removed_object_identifiers,
        budget,
    )?;
    let changed_members = member_edits.member_names().map(str::to_owned).collect();
    Ok(NativePopUpOutput {
        member_edits,
        cell_bytes: transition.cell_bytes,
        format_identifier: transition.format_identifier,
        control_cell_spec_identifier: transition.control_cell_spec_identifier,
        popup_model_identifier: transition.popup_model_identifier,
        added_object_identifiers: transition.added_object_identifiers,
        removed_object_identifiers: transition.removed_object_identifiers,
        changed_member_names: changed_members,
    })
}

/// Return the current popup model IDs reachable from a split control graph.
/// This mirrors [`existing_popup_model_identifiers`] but first builds the same
/// private merged view used by the write path, so the facade can preflight UUID
/// ownership without silently ignoring a sidecar member.
pub(super) fn existing_popup_model_identifiers_multi(
    input: NativePopUpGraphInput<'_>,
    budget: &mut TransactionBudget,
) -> Result<Vec<u64>> {
    let source_object_count = input
        .members
        .iter()
        .try_fold(0usize, |total, member| {
            total.checked_add(member.archive.objects.len())
        })
        .ok_or(NativePopUpError::InvalidSource)?;
    let source_payload_bytes = input
        .members
        .iter()
        .flat_map(|member| member.archive.objects.iter())
        .flat_map(|object| object.messages.iter())
        .try_fold(0usize, |total, message| {
            total.checked_add(message.data.len())
        })
        .ok_or(NativePopUpError::InvalidSource)?;
    let merge_allocation_count = source_object_count
        .checked_mul(2)
        .and_then(|count| count.checked_add(input.members.len()))
        .ok_or(NativePopUpError::InvalidSource)?;
    budget
        .charge_allocations(merge_allocation_count, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    budget
        .charge_scratch_bytes(source_payload_bytes, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    let (working, _) = merge_popup_members(&input)?;
    let placement = input
        .members
        .get(input.creation_member_index)
        .ok_or(NativePopUpError::InvalidSource)?;
    let synthetic = NativePopUpInput {
        archive: &working,
        component_index: placement.component_index,
        model_identifier: input.model.identifier,
        tile_identifier: input.tile.identifier,
        tile_row: input.tile_row,
        tile_column: input.tile_column,
        control_table_identifier: input.control_table_identifier,
        format_table_identifier: input.format_table_identifier,
        member_name: placement.member_name,
        format_payload: input.format_payload,
        new_popup_model_identifier: input.new_popup_model_identifier,
        desired: input.desired,
        limits: input.limits,
        path: input.path,
    };
    existing_popup_model_identifiers(synthetic, budget, input.path)
}

fn merge_popup_members(input: &NativePopUpGraphInput<'_>) -> Result<(Archive, Vec<(u64, usize)>)> {
    if input.members.is_empty() {
        return Err(NativePopUpError::InvalidSource);
    }
    if input.creation_member_index >= input.members.len()
        || input.model.member_index >= input.members.len()
        || input.tile.member_index >= input.members.len()
    {
        return Err(NativePopUpError::InvalidSource);
    }
    let mut working = Archive::new();
    let object_count = input
        .members
        .iter()
        .try_fold(0usize, |total, member| {
            total.checked_add(member.archive.objects.len())
        })
        .ok_or(NativePopUpError::InvalidSource)?;
    working
        .objects
        .try_reserve_exact(object_count)
        .map_err(|_| NativePopUpError::Allocation)?;
    let mut ownership = Vec::new();
    ownership
        .try_reserve_exact(object_count)
        .map_err(|_| NativePopUpError::Allocation)?;
    for (member_index, member) in input.members.iter().enumerate() {
        if member.member_name.is_empty()
            || input.members[..member_index].iter().any(|prior| {
                prior.component_index == member.component_index
                    || prior.member_name == member.member_name
            })
        {
            return Err(NativePopUpError::UnsupportedDependency);
        }
        for object in &member.archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or(NativePopUpError::InvalidSource)?;
            if identifier == 0 || ownership.iter().any(|(id, _)| *id == identifier) {
                return Err(NativePopUpError::UnsupportedDependency);
            }
            ownership.push((identifier, member_index));
            working.objects.push(object.clone());
        }
    }
    Ok((working, ownership))
}

fn split_popup_candidate(
    input: &NativePopUpGraphInput<'_>,
    ownership: &[(u64, usize)],
    candidate: &Archive,
    added: &[u64],
    removed: &[u64],
    budget: &mut TransactionBudget,
) -> Result<NativeControlOutput> {
    let mut per_member: Vec<Archive> = Vec::new();
    per_member
        .try_reserve_exact(input.members.len())
        .map_err(|_| NativePopUpError::Allocation)?;
    for member in input.members {
        let mut archive = Archive::new();
        archive
            .objects
            .try_reserve_exact(member.archive.objects.len())
            .map_err(|_| NativePopUpError::Allocation)?;
        per_member.push(archive);
    }
    let mut seen = Vec::new();
    seen.try_reserve_exact(candidate.objects.len())
        .map_err(|_| NativePopUpError::Allocation)?;
    for object in &candidate.objects {
        let identifier = object
            .archive_info
            .identifier
            .ok_or(NativePopUpError::InvalidSource)?;
        if seen.contains(&identifier) {
            return Err(NativePopUpError::UnsupportedDependency);
        }
        seen.push(identifier);
        let member_index =
            if let Some((_, index)) = ownership.iter().find(|(id, _)| *id == identifier) {
                *index
            } else if added.contains(&identifier) {
                input.creation_member_index
            } else {
                return Err(NativePopUpError::InvalidSource);
            };
        per_member
            .get_mut(member_index)
            .ok_or(NativePopUpError::InvalidSource)?
            .objects
            .push(object.clone());
    }
    for &(identifier, member_index) in ownership {
        let source_member = input
            .members
            .get(member_index)
            .ok_or(NativePopUpError::InvalidSource)?;
        let source_object = source_member
            .archive
            .objects
            .iter()
            .find(|object| object.archive_info.identifier == Some(identifier))
            .ok_or(NativePopUpError::InvalidSource)?;
        let present = per_member[member_index]
            .objects
            .iter()
            .any(|object| object.archive_info.identifier == Some(identifier));
        if !present && !removed.contains(&identifier) {
            return Err(NativePopUpError::InvalidSource);
        }
        if present {
            let candidate_object = per_member[member_index]
                .objects
                .iter()
                .find(|object| object.archive_info.identifier == Some(identifier))
                .ok_or(NativePopUpError::InvalidSource)?;
            if !source_object.same_content_ignoring_offsets(candidate_object)
                && removed.contains(&identifier)
            {
                return Err(NativePopUpError::InvalidSource);
            }
        }
    }
    // Removal compacts the object vector, so comparing by absolute position
    // would incorrectly reject every object following a culled popup.  The
    // source order must instead equal source IDs with removals filtered,
    // followed by any newly-created object assigned to this member.
    for (member_index, member) in input.members.iter().enumerate() {
        let mut expected = member
            .archive
            .objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .filter(|identifier| !removed.contains(identifier))
            .collect::<Vec<_>>();
        expected.extend(
            added
                .iter()
                .copied()
                .filter(|_identifier| input.creation_member_index == member_index),
        );
        let actual = per_member[member_index]
            .objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .collect::<Vec<_>>();
        if expected != actual {
            return Err(NativePopUpError::InvalidSource);
        }
    }
    // The exact-byte publication filter below retains one candidate and one
    // source serialization at the same time. Preflight both temporary Vecs
    // for every member before the first serialization allocation; charging
    // all members is conservative when the structural comparison later
    // proves some of them unchanged.
    let comparison_allocations = input
        .members
        .len()
        .checked_mul(2)
        .and_then(|count| count.checked_add(1))
        .ok_or(NativePopUpError::InvalidSource)?;
    let comparison_bytes =
        input
            .members
            .iter()
            .zip(&per_member)
            .try_fold(0usize, |total, (source, candidate)| {
                let source_len = source
                    .archive
                    .encoded_len_with_limits(input.limits)
                    .map_err(|_| NativePopUpError::Archive)?;
                let candidate_len = candidate
                    .encoded_len_with_limits(input.limits)
                    .map_err(|_| NativePopUpError::Archive)?;
                total
                    .checked_add(source_len)
                    .and_then(|value| value.checked_add(candidate_len))
                    .ok_or(NativePopUpError::InvalidSource)
            })?;
    budget
        .charge_allocations(comparison_allocations, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    budget
        .charge_scratch_bytes(comparison_bytes, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    budget
        .charge_transaction_work(comparison_bytes, input.path)
        .map_err(|_| NativePopUpError::Limit)?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(input.members.len())
        .map_err(|_| NativePopUpError::Allocation)?;
    for (member_index, member) in input.members.iter().enumerate() {
        let source_objects = &member.archive.objects;
        let candidate_objects = &per_member[member_index].objects;
        let source_ids = source_objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .filter(|identifier| !removed.contains(identifier))
            .collect::<Vec<_>>();
        let candidate_ids = candidate_objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .collect::<Vec<_>>();
        let changed = source_ids != candidate_ids
            || source_objects.iter().any(|source_object| {
                let Some(identifier) = source_object.archive_info.identifier else {
                    return true;
                };
                if removed.contains(&identifier) {
                    return true;
                }
                let Some(candidate_object) = candidate_objects
                    .iter()
                    .find(|object| object.archive_info.identifier == Some(identifier))
                else {
                    return true;
                };
                !source_object.same_content_ignoring_offsets(candidate_object)
            });
        if !changed {
            continue;
        }
        // Recheck the preflighted encoded lengths immediately before each
        // serialization.  `to_bytes_with_limits` derives its output length
        // from the current object metadata, so retaining these values here
        // proves that the two temporary Vecs below are still covered by the
        // aggregate preflight and that no late header-length change can evade
        // the source/candidate scratch and work charges.
        let expected_candidate_len = per_member[member_index]
            .encoded_len_with_limits(input.limits)
            .map_err(|_| NativePopUpError::Archive)?;
        let expected_source_len = member
            .archive
            .encoded_len_with_limits(input.limits)
            .map_err(|_| NativePopUpError::Archive)?;
        let bytes = per_member[member_index]
            .to_bytes_with_limits(input.limits)
            .map_err(|_| NativePopUpError::Archive)?;
        if bytes.len() != expected_candidate_len {
            return Err(NativePopUpError::InvalidSource);
        }
        // Object provenance can differ after the private merge/split even
        // when the physical member is byte-identical (for example, an
        // untouched popup object retains the source payload but receives a
        // fresh in-memory archive position).  Compare the serialized source
        // member before publishing an edit; metadata token selectors must
        // describe actual byte changes, never merely private provenance.
        let source_bytes = member
            .archive
            .to_bytes_with_limits(input.limits)
            .map_err(|_| NativePopUpError::Archive)?;
        if source_bytes.len() != expected_source_len {
            return Err(NativePopUpError::InvalidSource);
        }
        if bytes == source_bytes {
            continue;
        }
        edits.push(NativeMemberEdit::new(
            member.component_index,
            member.member_name,
            bytes,
        ));
    }
    NativeControlOutput::from_edits(edits)
}

/// Census the popup model IDs reachable from the selected control list
/// without cloning or rewriting the archive.  The facade runs this pass
/// before native candidate allocation so metadata can prove exact current
/// UUID ownership for every reused model.
pub(super) fn existing_popup_model_identifiers(
    input: NativePopUpInput<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u64>> {
    validate_unique_route(input.archive, input.control_table_identifier)?;
    let control_object = unique_object(input.archive, input.control_table_identifier)?;
    let options = control_object
        .messages
        .iter()
        .filter(|message| message.type_ == TABLE_DATA_LIST_TYPE)
        .max_by_key(|message| message.data.len())
        .map(|message| budget.residual_storage_options(&message.data))
        .ok_or(NativePopUpError::InvalidSource)?;
    let control = unique_list_message(
        control_object,
        LIST_CONTROL_CELL_SPEC,
        options,
        budget,
        path,
    )?;
    let identifiers = popup_references_from_list(control.payload)?;
    for &identifier in &identifiers {
        let popup = unique_object(input.archive, identifier)?;
        let message_index = unique_message_index(popup, POPUP_MODEL_TYPE)?;
        let (_, report) = popup_codec::decode_popup_menu_model_with_report(
            &popup.messages[message_index].data,
            budget.residual_popup_options(&popup.messages[message_index].data),
        )
        .map_err(|_| NativePopUpError::Codec)?;
        budget
            .charge_wire_bytes(report.input_bytes(), path)
            .and_then(|_| budget.charge_wire_fields(report.fields(), path))
            .and_then(|_| budget.charge_wire_work(report.work_bytes(), path))
            .and_then(|_| budget.charge_wire_nesting(report.max_depth(), path))
            .and_then(|_| budget.charge_payload_references(report.references(), path))
            .and_then(|_| budget.charge_payload_items(report.items(), path))
            .and_then(|_| budget.charge_wire_text_bytes(report.text_bytes(), path))
            .and_then(|_| budget.charge_allocations(report.allocations(), path))
            .and_then(|_| budget.charge_scratch_bytes(report.scratch_bytes(), path))
            .and_then(|_| budget.charge_retained_bytes(report.retained_bytes(), path))
            .map_err(|_| NativePopUpError::Limit)?;
    }
    Ok(identifiers)
}

#[derive(Debug, Clone, Copy)]
struct ListMessage<'source> {
    message_index: usize,
    payload: &'source [u8],
}

#[derive(Debug, Clone, Copy)]
struct OldState {
    format_key: u32,
    control_key: u32,
    popup_identifier: Option<u64>,
}

#[derive(Debug, Clone)]
struct DesiredState {
    format_key: u32,
    control_key: u32,
    popup_identifier: Option<u64>,
    starts_with_first: bool,
    first_item_string_identifier: Option<u32>,
    new_popup_payload: Option<Vec<u8>>,
}

fn prepare_desired_state(
    archive: &Archive,
    control_payload: &[u8],
    desired: NativePopUpValue<'_>,
    rooted_popup_identifiers: &[u64],
    new_identifier: Option<u64>,
    options: storage_codec::DecodeOptions,
    budget: &mut TransactionBudget,
) -> Result<DesiredState> {
    if desired.items.is_empty() {
        return Err(NativePopUpError::InvalidSource);
    }
    let popup_options = popup_options_for_items(desired.items);
    let prepared = popup_codec::prepare_popup_menu_model_write(desired.items, popup_options)
        .map_err(|_| NativePopUpError::Codec)?;
    let requirements = prepared.execution_requirements();
    charge_popup_write_requirements(budget, requirements, Path::Package)?;
    let canonical = prepared
        .execute(popup_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|_| NativePopUpError::Codec)?
        .into_bytes();
    let mut popup_identifier = None;
    let mut new_popup_payload = None;
    for object in &archive.objects {
        if object.archive_info.identifier == Some(0) {
            return Err(NativePopUpError::InvalidSource);
        }
        let mut messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == POPUP_MODEL_TYPE);
        let Some(message) = messages.next() else {
            continue;
        };
        if messages.next().is_some() {
            return Err(NativePopUpError::InvalidSource);
        }
        let object_identifier = object
            .archive_info
            .identifier
            .ok_or(NativePopUpError::InvalidSource)?;
        if !rooted_popup_identifiers.contains(&object_identifier) {
            continue;
        }
        let (snapshot, report) = popup_codec::decode_popup_menu_model_with_report(
            &message.data,
            budget.residual_popup_options(&message.data),
        )
        .map_err(|_| NativePopUpError::Codec)?;
        budget
            .charge_wire_bytes(report.input_bytes(), Path::Package)
            .and_then(|_| budget.charge_wire_fields(report.fields(), Path::Package))
            .and_then(|_| budget.charge_wire_work(report.work_bytes(), Path::Package))
            .and_then(|_| budget.charge_wire_nesting(report.max_depth(), Path::Package))
            .and_then(|_| budget.charge_payload_references(report.references(), Path::Package))
            .and_then(|_| budget.charge_payload_items(report.items(), Path::Package))
            .and_then(|_| budget.charge_wire_text_bytes(report.text_bytes(), Path::Package))
            .and_then(|_| budget.charge_allocations(report.allocations(), Path::Package))
            .and_then(|_| budget.charge_scratch_bytes(report.scratch_bytes(), Path::Package))
            .and_then(|_| budget.charge_retained_bytes(report.retained_bytes(), Path::Package))
            .map_err(|_| NativePopUpError::Limit)?;
        if snapshot
            .items()
            .map(|item| item.value())
            .eq(desired.items.iter().copied())
        {
            if popup_identifier.replace(object_identifier).is_some() {
                return Err(NativePopUpError::UnsupportedDependency);
            }
        }
    }
    if popup_identifier.is_none() {
        let id = new_identifier.ok_or(NativePopUpError::UnsupportedDependency)?;
        if id == 0 || archive.object(id).is_some() {
            return Err(NativePopUpError::InvalidSource);
        }
        popup_identifier = Some(id);
        new_popup_payload = Some(canonical);
    }
    let popup_identifier = popup_identifier.ok_or(NativePopUpError::InvalidSource)?;
    let control = decode_list_entries(control_payload, options)?;
    let mut control_key = None;
    for entry in &control.entries {
        if entry.payload_kind != PayloadKind::ControlCellSpec {
            continue;
        }
        if entry.ref_count == 0 {
            return Err(NativePopUpError::InvalidSource);
        }
        let (snapshot, report) = control_codec::decode_any_cell_spec_with_report(
            &entry.payload,
            budget.residual_popup_options(&entry.payload),
        )
        .map_err(|_| NativePopUpError::Codec)?;
        budget
            .charge_wire_bytes(report.input_bytes(), Path::Package)
            .and_then(|_| budget.charge_wire_fields(report.fields(), Path::Package))
            .and_then(|_| budget.charge_wire_work(report.work_bytes(), Path::Package))
            .and_then(|_| budget.charge_wire_nesting(report.max_depth(), Path::Package))
            .and_then(|_| budget.charge_payload_references(report.references(), Path::Package))
            .and_then(|_| budget.charge_payload_items(report.items(), Path::Package))
            .and_then(|_| budget.charge_wire_text_bytes(report.text_bytes(), Path::Package))
            .and_then(|_| budget.charge_allocations(report.allocations(), Path::Package))
            .and_then(|_| budget.charge_scratch_bytes(report.scratch_bytes(), Path::Package))
            .and_then(|_| budget.charge_retained_bytes(report.retained_bytes(), Path::Package))
            .map_err(|_| NativePopUpError::Limit)?;
        if let control_codec::CellSpecSnapshot::Popup(snapshot) = snapshot
            && snapshot.popup_model().identifier() == popup_identifier
            && snapshot.starts_with_first() == desired.starts_with_first
        {
            if control_key.replace(entry.key).is_some() {
                return Err(NativePopUpError::UnsupportedDependency);
            }
        }
    }
    if control_key.is_none() {
        let mut candidate = control.next_list_id.max(1);
        while control.entries.iter().any(|entry| entry.key == candidate) {
            candidate = candidate
                .checked_add(1)
                .ok_or(NativePopUpError::InvalidSource)?;
        }
        control_key = Some(candidate);
    }
    let control_key = control_key.ok_or(NativePopUpError::InvalidSource)?;
    Ok(DesiredState {
        format_key: 0,
        control_key,
        popup_identifier: Some(popup_identifier),
        starts_with_first: desired.starts_with_first,
        first_item_string_identifier: desired.first_item_string_identifier,
        new_popup_payload,
    })
}

fn validate_control_list_payloads(
    source: &[u8],
    options: storage_codec::DecodeOptions,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<()> {
    let list = decode_list_entries(source, options)?;
    for entry in list.entries {
        if entry.payload_kind != PayloadKind::ControlCellSpec {
            continue;
        }
        if entry.ref_count == 0 {
            return Err(NativePopUpError::InvalidSource);
        }
        let (_, report) = control_codec::decode_any_cell_spec_with_report(
            &entry.payload,
            budget.residual_popup_options(&entry.payload),
        )
        .map_err(|_| NativePopUpError::Codec)?;
        budget
            .charge_wire_bytes(report.input_bytes(), path)
            .and_then(|_| budget.charge_wire_fields(report.fields(), path))
            .and_then(|_| budget.charge_wire_work(report.work_bytes(), path))
            .and_then(|_| budget.charge_wire_nesting(report.max_depth(), path))
            .and_then(|_| budget.charge_payload_references(report.references(), path))
            .and_then(|_| budget.charge_payload_items(report.items(), path))
            .and_then(|_| budget.charge_wire_text_bytes(report.text_bytes(), path))
            .and_then(|_| budget.charge_allocations(report.allocations(), path))
            .and_then(|_| budget.charge_scratch_bytes(report.scratch_bytes(), path))
            .and_then(|_| budget.charge_retained_bytes(report.retained_bytes(), path))
            .map_err(|_| NativePopUpError::Limit)?;
    }
    Ok(())
}

#[derive(Debug, Default)]
struct BncReferenceCounts {
    format: Vec<(u32, u32)>,
    control: Vec<(u32, u32)>,
}

impl BncReferenceCounts {
    fn increment(entries: &mut Vec<(u32, u32)>, identifier: u32) -> Result<()> {
        if let Some((_, count)) = entries.iter_mut().find(|(key, _)| *key == identifier) {
            *count = count
                .checked_add(1)
                .ok_or(NativePopUpError::InvalidSource)?;
            return Ok(());
        }
        entries
            .try_reserve_exact(1)
            .map_err(|_| NativePopUpError::Allocation)?;
        entries.push((identifier, 1));
        Ok(())
    }

    fn count(entries: &[(u32, u32)], identifier: u32) -> u32 {
        entries
            .iter()
            .find(|(key, _)| *key == identifier)
            .map_or(0, |(_, count)| *count)
    }
}

struct BncReferenceVisitor<'a> {
    counts: &'a mut BncReferenceCounts,
    failed: bool,
}

impl storage_codec::StorageVisitor for BncReferenceVisitor<'_> {
    fn visit_tile_row(
        &mut self,
        row: storage_codec::TileRowInfoSnapshot<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        if census_bnc_row(row, self.counts).is_err() {
            self.failed = true;
        }
        Ok(())
    }
}

fn census_bnc_references(
    archive: &Archive,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<BncReferenceCounts> {
    let mut counts = BncReferenceCounts::default();
    for object in &archive.objects {
        let mut messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == TILE_TYPE);
        let Some(message) = messages.next() else {
            continue;
        };
        if messages.next().is_some() {
            return Err(NativePopUpError::InvalidSource);
        }
        let mut visitor = BncReferenceVisitor {
            counts: &mut counts,
            failed: false,
        };
        let (_, report) = storage_codec::decode_tile_with_visitor(
            &message.data,
            budget.residual_storage_options(&message.data),
            &mut visitor,
        )
        .map_err(|_| NativePopUpError::Codec)?;
        budget
            .charge_storage_decode_report(report, true, path)
            .map_err(|_| NativePopUpError::Limit)?;
        if visitor.failed {
            return Err(NativePopUpError::InvalidSource);
        }
    }
    Ok(counts)
}

fn census_bnc_row(
    row: storage_codec::TileRowInfoSnapshot<'_>,
    counts: &mut BncReferenceCounts,
) -> Result<()> {
    let (storage, offsets) = match (row.cell_storage_buffer(), row.cell_offsets()) {
        (Some(storage), Some(offsets)) => (storage, offsets),
        (None, None) => (
            row.cell_storage_buffer_pre_bnc(),
            row.cell_offsets_pre_bnc(),
        ),
        _ => return Err(NativePopUpError::InvalidSource),
    };
    if !offsets.len().is_multiple_of(2) {
        return Err(NativePopUpError::InvalidSource);
    }
    let unit = if row.has_wide_offsets().unwrap_or(false) {
        4usize
    } else {
        1usize
    };
    let mut previous = None;
    let mut occupied = 0usize;
    for encoded in offsets.chunks_exact(2) {
        let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
        if raw == u16::MAX {
            continue;
        }
        let start = usize::from(raw)
            .checked_mul(unit)
            .ok_or(NativePopUpError::InvalidSource)?;
        // A minimal automatic BNC cell has an empty payload, so adjacent
        // materialized columns may legitimately share the same start offset.
        // Decreasing offsets remain invalid; equal offsets describe a
        // zero-length cell and are parsed through `BncCell::parse(&[])`.
        if start > storage.len() || previous.is_some_and(|prior| prior > start) {
            return Err(NativePopUpError::InvalidSource);
        }
        if let Some(prior) = previous {
            census_bnc_cell(&storage[prior..start], counts)?;
        }
        previous = Some(start);
        occupied = occupied
            .checked_add(1)
            .ok_or(NativePopUpError::InvalidSource)?;
    }
    if let Some(start) = previous {
        census_bnc_cell(&storage[start..], counts)?;
    }
    if occupied != usize::try_from(row.cell_count()).map_err(|_| NativePopUpError::InvalidSource)? {
        return Err(NativePopUpError::InvalidSource);
    }
    Ok(())
}

fn census_bnc_cell(bytes: &[u8], counts: &mut BncReferenceCounts) -> Result<()> {
    let cell = BncCell::parse(bytes).map_err(|_| NativePopUpError::InvalidSource)?;
    match (
        cell.format_identifier(),
        cell.control_cell_spec_identifier(),
    ) {
        (None, None) => Ok(()),
        (Some(format), Some(control)) => {
            if format == 0 || control == 0 {
                return Err(NativePopUpError::InvalidSource);
            }
            BncReferenceCounts::increment(&mut counts.format, format)?;
            BncReferenceCounts::increment(&mut counts.control, control)
        },
        _ => Err(NativePopUpError::InvalidSource),
    }
}

fn validate_bnc_refcounts(
    counts: &BncReferenceCounts,
    payload: &[u8],
    list_type: i32,
    options: storage_codec::DecodeOptions,
) -> Result<()> {
    let list = decode_list_entries(payload, options)?;
    if list_type == LIST_CONTROL_CELL_SPEC {
        for (index, entry) in list.entries.iter().enumerate() {
            if list.entries[index + 1..].iter().any(|other| {
                entry.payload_kind == PayloadKind::ControlCellSpec
                    && other.payload_kind == PayloadKind::ControlCellSpec
                    && entry.payload == other.payload
            }) {
                // Equal display-format payloads are valid in native format
                // lists: distinct scalar controls can share the same
                // formatting while retaining separate keys.  Control-list
                // entries are different: two byte-identical CellSpecs make
                // popup reuse/refcount ownership ambiguous, so require one
                // canonical entry for an equal control payload.
                return Err(NativePopUpError::UnsupportedDependency);
            }
        }
    }
    let mut seen = Vec::new();
    for entry in list.entries {
        let observed = match list_type {
            LIST_FORMAT if entry.payload_kind == PayloadKind::Format => {
                BncReferenceCounts::count(&counts.format, entry.key)
            },
            LIST_CONTROL_CELL_SPEC if entry.payload_kind == PayloadKind::ControlCellSpec => {
                BncReferenceCounts::count(&counts.control, entry.key)
            },
            _ => continue,
        };
        if observed != entry.ref_count {
            return Err(NativePopUpError::InvalidSource);
        }
        if list_type == LIST_FORMAT && entry.payload_kind == PayloadKind::Format
            || list_type == LIST_CONTROL_CELL_SPEC
                && entry.payload_kind == PayloadKind::ControlCellSpec
        {
            seen.push(entry.key);
        }
    }
    let expected = if list_type == LIST_FORMAT {
        &counts.format
    } else {
        &counts.control
    };
    if expected
        .iter()
        .any(|(identifier, _)| !seen.contains(identifier))
    {
        return Err(NativePopUpError::InvalidSource);
    }
    Ok(())
}

fn inspect_old_state(
    archive: &Archive,
    control_payload: &[u8],
    old_control: Option<u32>,
    old_format: Option<u32>,
    options: storage_codec::DecodeOptions,
) -> Result<Option<OldState>> {
    let (Some(control_key), Some(format_key)) = (old_control, old_format) else {
        if old_control.is_some() || old_format.is_some() {
            return Err(NativePopUpError::InvalidSource);
        }
        return Ok(None);
    };
    let list = decode_list_entries(control_payload, options)?;
    let entry = list
        .entries
        .iter()
        .find(|entry| entry.key == control_key)
        .ok_or(NativePopUpError::InvalidSource)?;
    if entry.ref_count == 0 || entry.payload_kind != PayloadKind::ControlCellSpec {
        return Err(NativePopUpError::InvalidSource);
    }
    let (spec, _) = popup_codec::decode_cell_spec_with_report(
        &entry.payload,
        popup_options_for_bytes(&entry.payload),
    )
    .map_err(|_| NativePopUpError::Codec)?;
    if spec.interaction_type() != 7
        || spec.popup_model().identifier() == 0
        || spec.starts_with_first() && spec.popup_model().identifier() == 0
    {
        return Err(NativePopUpError::InvalidSource);
    }
    let popup_identifier = spec.popup_model().identifier();
    let popup = unique_object(archive, popup_identifier)?;
    let popup_message = unique_message_index(popup, POPUP_MODEL_TYPE)?;
    popup_codec::decode_popup_menu_model_with_report(
        &popup.messages[popup_message].data,
        popup_options_for_bytes(&popup.messages[popup_message].data),
    )
    .map_err(|_| NativePopUpError::Codec)?;
    Ok(Some(OldState {
        format_key,
        control_key,
        popup_identifier: Some(popup_identifier),
    }))
}

fn rewrite_list_for_format(
    source: &[u8],
    old_key: Option<u32>,
    _desired_key: Option<u32>,
    payload: Option<&[u8]>,
    options: storage_codec::DecodeOptions,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(Vec<u8>, Option<u32>)> {
    let list = decode_list_entries(source, options)?;
    let Some(payload) = payload else {
        let Some(old_key) = old_key else {
            return Ok((source.to_owned(), None));
        };
        let entry = list
            .entries
            .iter()
            .find(|entry| entry.key == old_key)
            .ok_or(NativePopUpError::InvalidSource)?;
        let mutation = if entry.ref_count > 1 {
            storage_codec::TableDataListEntryMutation::RefCount(
                storage_codec::TableDataListEntryRefCountEdit::new(
                    old_key,
                    entry.ref_count,
                    entry.ref_count - 1,
                ),
            )
        } else {
            storage_codec::TableDataListEntryMutation::Remove(
                storage_codec::TableDataListEntryRemovalSpec::format(
                    old_key,
                    entry.ref_count,
                    &entry.payload,
                ),
            )
        };
        return Ok((
            apply_list_mutation_with_budget(source, mutation, budget, path)?,
            None,
        ));
    };
    let mut desired_key = list
        .entries
        .iter()
        .find(|entry| entry.payload_kind == PayloadKind::Format && entry.payload == payload)
        .map(|entry| entry.key)
        .unwrap_or_else(|| list.next_list_id.max(1));
    while list.entries.iter().any(|entry| {
        entry.key == desired_key
            && !(entry.payload_kind == PayloadKind::Format && entry.payload == payload)
    }) {
        desired_key = desired_key
            .checked_add(1)
            .ok_or(NativePopUpError::InvalidSource)?;
    }
    let mut output = source.to_owned();
    if let Some(old_key) = old_key
        && old_key != desired_key
        && let Some(entry) = list.entries.iter().find(|entry| entry.key == old_key)
    {
        let mutation = if entry.ref_count > 1 {
            storage_codec::TableDataListEntryMutation::RefCount(
                storage_codec::TableDataListEntryRefCountEdit::new(
                    old_key,
                    entry.ref_count,
                    entry.ref_count - 1,
                ),
            )
        } else {
            storage_codec::TableDataListEntryMutation::Remove(
                storage_codec::TableDataListEntryRemovalSpec::format(
                    old_key,
                    entry.ref_count,
                    &entry.payload,
                ),
            )
        };
        output = apply_list_mutation_with_budget(&output, mutation, budget, path)?;
    }
    if list.entries.iter().any(|entry| entry.key == desired_key) {
        if old_key != Some(desired_key) {
            let entry = list
                .entries
                .iter()
                .find(|entry| entry.key == desired_key)
                .ok_or(NativePopUpError::InvalidSource)?;
            output = apply_list_mutation_with_budget(
                &output,
                storage_codec::TableDataListEntryMutation::RefCount(
                    storage_codec::TableDataListEntryRefCountEdit::new(
                        desired_key,
                        entry.ref_count,
                        entry.ref_count.saturating_add(1),
                    ),
                ),
                budget,
                path,
            )?;
        }
    } else {
        output = apply_list_mutation_with_budget(
            &output,
            storage_codec::TableDataListEntryMutation::Append(
                storage_codec::TableDataListEntryAppend::format(desired_key, 1, payload),
            ),
            budget,
            path,
        )?;
    }
    Ok((output, Some(desired_key)))
}

fn rewrite_control_list(
    source: &[u8],
    old: Option<OldState>,
    desired: Option<&DesiredState>,
    options: storage_codec::DecodeOptions,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(Vec<u8>, Option<u32>, Option<u64>)> {
    let original = decode_list_entries(source, options)?;
    let mut output = source.to_owned();
    let desired_key = desired.as_ref().map(|state| state.control_key);
    let desired_popup = desired.as_ref().and_then(|state| state.popup_identifier);
    if let Some(old) = old {
        if desired_key != Some(old.control_key) {
            let entry = original
                .entries
                .iter()
                .find(|entry| entry.key == old.control_key)
                .ok_or(NativePopUpError::InvalidSource)?;
            let mutation = if entry.ref_count > 1 {
                storage_codec::TableDataListEntryMutation::RefCount(
                    storage_codec::TableDataListEntryRefCountEdit::new(
                        old.control_key,
                        entry.ref_count,
                        entry.ref_count - 1,
                    ),
                )
            } else {
                storage_codec::TableDataListEntryMutation::Remove(
                    storage_codec::TableDataListEntryRemovalSpec::control_cell_spec(
                        old.control_key,
                        entry.ref_count,
                        &entry.payload,
                    ),
                )
            };
            output = apply_list_mutation_with_budget(&output, mutation, budget, path)?;
        }
    }
    if let Some(desired) = desired {
        let after = decode_list_entries(&output, options)?;
        if let Some(entry) = after
            .entries
            .iter()
            .find(|entry| entry.key == desired.control_key)
        {
            if old.map(|old| old.control_key) != Some(desired.control_key) {
                output = apply_list_mutation_with_budget(
                    &output,
                    storage_codec::TableDataListEntryMutation::RefCount(
                        storage_codec::TableDataListEntryRefCountEdit::new(
                            desired.control_key,
                            entry.ref_count,
                            entry.ref_count.saturating_add(1),
                        ),
                    ),
                    budget,
                    path,
                )?;
            }
        } else {
            let popup_identifier = desired_popup.ok_or(NativePopUpError::InvalidSource)?;
            let prepared = popup_codec::prepare_cell_spec_write(
                popup_identifier,
                desired.starts_with_first,
                popup_options_for_bytes(source),
            )
            .map_err(|_| NativePopUpError::Codec)?;
            let requirements = prepared.execution_requirements();
            charge_popup_write_requirements(budget, requirements, path)?;
            let spec = prepared
                .execute(popup_codec::RewriteExecutionLimits::exact(requirements))
                .map_err(|_| NativePopUpError::Codec)?
                .into_bytes();
            output = apply_list_mutation_with_budget(
                &output,
                storage_codec::TableDataListEntryMutation::Append(
                    storage_codec::TableDataListEntryAppend::control_cell_spec(
                        desired.control_key,
                        1,
                        &spec,
                    ),
                ),
                budget,
                path,
            )?;
        }
    }
    Ok((output, desired_key, desired_popup))
}

fn apply_list_mutation_with_budget(
    source: &[u8],
    mutation: storage_codec::TableDataListEntryMutation<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>> {
    let appended_key = match mutation {
        storage_codec::TableDataListEntryMutation::Append(append) => Some(append.key()),
        _ => None,
    };
    let options = budget.residual_storage_rewrite_options(source);
    let plan = storage_codec::prepare_table_data_list_entry_rewrite(source, mutation, options)
        .map_err(|_| NativePopUpError::Codec)?;
    let preparation = plan.prepare_report();
    // The prepared requirements already aggregate source, payload, result,
    // and verification fields/work.  Charge only source-only byte classes
    // from the preparation report here; charging its fields/work/references
    // as well would debit the same scan twice.
    budget
        .charge_wire_bytes(preparation.source_bytes(), path)
        .and_then(|_| budget.charge_wire_reference_bytes(preparation.reference_bytes(), path))
        .and_then(|_| budget.charge_wire_text_bytes(preparation.text_bytes(), path))
        .map_err(|_| NativePopUpError::Limit)?;
    let requirements = plan.requirements();
    budget
        .charge_wire_fields(requirements.fields(), path)
        .and_then(|_| budget.charge_wire_work(requirements.work_bytes(), path))
        .and_then(|_| budget.charge_wire_nesting(requirements.max_depth(), path))
        .and_then(|_| budget.charge_payload_references(requirements.references(), path))
        .and_then(|_| budget.charge_scratch_bytes(requirements.scratch_bytes(), path))
        .and_then(|_| budget.charge_retained_bytes(requirements.retained_bytes(), path))
        .and_then(|_| budget.charge_allocations(requirements.allocations(), path))
        .and_then(|_| budget.charge_transaction_work(requirements.output_bytes(), path))
        .map_err(|_| NativePopUpError::Limit)?;
    let (bytes, report) = plan
        .execute(requirements.exact_limits())
        .map_err(|_| NativePopUpError::Codec)?;
    if report.output_bytes() != requirements.output_bytes()
        || report.fields() != requirements.fields()
        || report.work_bytes() != requirements.work_bytes()
        || report.references() != requirements.references()
        || report.scratch_bytes() != requirements.scratch_bytes()
        || report.allocations() != requirements.allocations()
        || report.retained_bytes() != requirements.retained_bytes()
    {
        return Err(NativePopUpError::InvalidSource);
    }
    match appended_key {
        Some(key) => advance_table_data_list_next_id(&bytes, key),
        None => Ok(bytes),
    }
}

fn charge_popup_write_requirements(
    budget: &mut TransactionBudget,
    requirements: popup_codec::RewriteExecutionRequirements,
    path: Path,
) -> Result<()> {
    budget
        .charge_wire_fields(requirements.fields(), path)
        .and_then(|_| budget.charge_wire_work(requirements.work_bytes(), path))
        .and_then(|_| budget.charge_wire_nesting(requirements.max_depth(), path))
        .and_then(|_| budget.charge_payload_references(requirements.references(), path))
        .and_then(|_| budget.charge_payload_items(requirements.items(), path))
        .and_then(|_| budget.charge_wire_text_bytes(requirements.text_bytes(), path))
        .and_then(|_| budget.charge_allocations(requirements.allocations(), path))
        .and_then(|_| budget.charge_scratch_bytes(requirements.scratch_bytes(), path))
        .and_then(|_| budget.charge_retained_bytes(requirements.retained_bytes(), path))
        .and_then(|_| budget.charge_transaction_work(requirements.output_bytes(), path))
        .map_err(|_| NativePopUpError::Limit)
}

fn table_data_list_type(
    source: &[u8],
    options: storage_codec::DecodeOptions,
) -> Result<Option<i32>> {
    storage_codec::decode_table_data_list_type_with_report(source, options)
        .map(|(snapshot, _)| Some(snapshot.list_type()))
        .map_err(|_| NativePopUpError::Codec)
}

fn advance_table_data_list_next_id(source: &[u8], appended_key: u32) -> Result<Vec<u8>> {
    let view = WireView::parse(source).map_err(|_| NativePopUpError::InvalidSource)?;
    let mut next = None;
    for field in view.fields() {
        if field.number() != 2 {
            continue;
        }
        if next.is_some() || field.wire_type() != 0 {
            return Err(NativePopUpError::InvalidSource);
        }
        let (value, consumed) = decode_varint_from_bytes(field.payload())
            .map_err(|_| NativePopUpError::InvalidSource)?;
        if consumed != field.payload().len() || value > u64::from(u32::MAX) {
            return Err(NativePopUpError::InvalidSource);
        }
        next = Some(value as u32);
    }
    let current = next.ok_or(NativePopUpError::InvalidSource)?;
    if appended_key < current {
        return Ok(source.to_owned());
    }
    let replacement = appended_key
        .checked_add(1)
        .ok_or(NativePopUpError::InvalidSource)?;
    let mut output = Vec::with_capacity(source.len());
    let mut replaced = false;
    for field in view.fields() {
        if field.number() == 2 {
            if replaced || field.wire_type() != 0 {
                return Err(NativePopUpError::InvalidSource);
            }
            litchi_iwa_common::wire::append_varint_field(&mut output, 2, u64::from(replacement))
                .map_err(|_| NativePopUpError::InvalidSource)?;
            replaced = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !replaced {
        return Err(NativePopUpError::InvalidSource);
    }
    Ok(output)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PayloadKind {
    Format,
    ControlCellSpec,
    Other,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ListEntry {
    key: u32,
    ref_count: u32,
    payload_kind: PayloadKind,
    payload: Vec<u8>,
}

#[derive(Debug)]
struct ListFacts {
    next_list_id: u32,
    entries: Vec<ListEntry>,
}

fn decode_list_entries(source: &[u8], options: storage_codec::DecodeOptions) -> Result<ListFacts> {
    decode_list_entries_with_report(source, options).map(|(facts, _)| facts)
}

fn decode_list_entries_with_report(
    source: &[u8],
    options: storage_codec::DecodeOptions,
) -> Result<(ListFacts, storage_codec::DecodeReport)> {
    let mut visitor = ListVisitor::default();
    let (list, report) =
        storage_codec::decode_table_data_list_with_visitor(source, options, &mut visitor)
            .map_err(|_| NativePopUpError::Codec)?;
    if list.list_type() != LIST_CONTROL_CELL_SPEC
        && list.list_type() != LIST_FORMAT
        && list.list_type() != LIST_STRING
    {
        return Err(NativePopUpError::InvalidSource);
    }
    if visitor.entries.iter().enumerate().any(|(index, entry)| {
        visitor.entries[index + 1..]
            .iter()
            .any(|other| other.key == entry.key)
    }) {
        return Err(NativePopUpError::InvalidSource);
    }
    // Segmented lists are a separate storage ownership route.  This first
    // popup transaction only rewrites one rooted list and must fail closed
    // rather than treating a segment reference as an opaque sidecar.
    if visitor.segments != 0 {
        return Err(NativePopUpError::UnsupportedDependency);
    }
    Ok((
        ListFacts {
            next_list_id: list.next_list_id(),
            entries: visitor.entries,
        },
        report,
    ))
}

#[derive(Default)]
struct ListVisitor {
    entries: Vec<ListEntry>,
    segments: usize,
}

impl storage_codec::StorageVisitor for ListVisitor {
    fn visit_list_entry_record(
        &mut self,
        record: storage_codec::TableDataListEntryRecord<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        let snapshot = record.snapshot();
        let (payload_kind, payload) = if let Some(payload) = snapshot.cell_spec() {
            (PayloadKind::ControlCellSpec, payload)
        } else if let Some(payload) = snapshot.format() {
            (PayloadKind::Format, payload)
        } else if let Some(value) = snapshot.string_value() {
            (PayloadKind::Other, value.as_bytes())
        } else {
            const EMPTY: &[u8] = &[];
            (PayloadKind::Other, EMPTY)
        };
        self.entries.push(ListEntry {
            key: snapshot.key(),
            ref_count: snapshot.ref_count(),
            payload_kind,
            payload: payload.to_vec(),
        });
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        _reference: storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        self.segments = self.segments.saturating_add(1);
        Ok(())
    }
}

fn unique_list_message<'source>(
    object: &'source ArchiveObject,
    list_type: i32,
    options: storage_codec::DecodeOptions,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<ListMessage<'source>> {
    let mut selected = None;
    for (message_index, message) in object.messages.iter().enumerate() {
        if message.type_ != TABLE_DATA_LIST_TYPE {
            continue;
        }
        // Route by the strict root envelope first so unrelated list schemas
        // remain opaque.  The selected list receives one full visitor pass;
        // unselected envelopes are charged only for this small route scan.
        let (snapshot, root_report) =
            storage_codec::decode_table_data_list_type_with_report(&message.data, options)
                .map_err(|_| NativePopUpError::Codec)?;
        if snapshot.list_type() != list_type {
            budget
                .charge_storage_decode_report(root_report, true, path)
                .map_err(|_| NativePopUpError::Limit)?;
            continue;
        }
        let (_, report) = decode_list_entries_with_report(&message.data, options)?;
        budget
            .charge_storage_decode_report(report, true, path)
            .map_err(|_| NativePopUpError::Limit)?;
        if selected.is_some() {
            return Err(NativePopUpError::InvalidSource);
        }
        let info = object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(NativePopUpError::InvalidSource)?;
        if info.object_references.contains(&0) || duplicate_identifiers(&info.object_references) {
            return Err(NativePopUpError::InvalidSource);
        }
        selected = Some(ListMessage {
            message_index,
            payload: &message.data,
        });
    }
    selected.ok_or(NativePopUpError::InvalidSource)
}

fn unique_object(archive: &Archive, identifier: u64) -> Result<&ArchiveObject> {
    if identifier == 0 {
        return Err(NativePopUpError::InvalidSource);
    }
    let mut matches = archive
        .objects
        .iter()
        .filter(|object| object.archive_info.identifier == Some(identifier));
    let object = matches.next().ok_or(NativePopUpError::InvalidSource)?;
    if matches.next().is_some() {
        return Err(NativePopUpError::InvalidSource);
    }
    Ok(object)
}

fn validate_unique_route(archive: &Archive, identifier: u64) -> Result<()> {
    let _ = unique_object(archive, identifier)?;
    Ok(())
}

fn unique_message_index(object: &ArchiveObject, message_type: u32) -> Result<usize> {
    let mut matches = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == message_type);
    let index = matches.next().map(|(index, _)| index);
    if matches.next().is_some() {
        return Err(NativePopUpError::InvalidSource);
    }
    index.ok_or(NativePopUpError::InvalidSource)
}

fn tile_identifier(
    source: &[u8],
    expected: u64,
    row: u32,
    options: storage_codec::DecodeOptions,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<u64> {
    let mut visitor = TileReferenceVisitor::default();
    let (storage, report) =
        storage_codec::decode_tile_storage_with_visitor(source, options, &mut visitor)
            .map_err(|_| NativePopUpError::Codec)?;
    budget
        .charge_storage_decode_report(report, true, path)
        .map_err(|_| NativePopUpError::Limit)?;
    let tile_size = storage.tile_size().ok_or(NativePopUpError::InvalidSource)?;
    let tile_id = row / tile_size.max(1);
    let mut matches = visitor
        .tiles
        .iter()
        .filter(|tile| tile.0 == tile_id)
        .map(|tile| tile.1);
    let identifier = matches.next().ok_or(NativePopUpError::InvalidSource)?;
    if matches.next().is_some() || identifier != expected {
        return Err(NativePopUpError::InvalidSource);
    }
    Ok(identifier)
}

#[derive(Default)]
struct TileReferenceVisitor {
    tiles: Vec<(u32, u64)>,
}

impl storage_codec::StorageVisitor for TileReferenceVisitor {
    fn visit_tile_reference(
        &mut self,
        record: storage_codec::TileReferenceRecord<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        self.tiles
            .push((record.tile_id(), record.reference().identifier()));
        Ok(())
    }
}

fn storage_options(source: &[u8]) -> storage_codec::DecodeOptions {
    let bytes = source.len().max(1);
    storage_codec::DecodeOptions::new(
        bytes,
        bytes.saturating_mul(16).max(1),
        bytes.saturating_mul(128).max(1),
        64,
        bytes.saturating_mul(8).max(1),
        bytes.saturating_mul(8).max(1),
    )
}

fn popup_options_for_bytes(source: &[u8]) -> popup_codec::DecodeOptions {
    popup_codec::DecodeOptions::for_source(source)
        .with_max_output_bytes(source.len().saturating_mul(4).max(1))
        .with_max_items(source.len().max(1))
        .with_max_text_bytes(source.len().saturating_mul(4).max(1))
}

fn popup_options_for_items(items: &[&str]) -> popup_codec::DecodeOptions {
    let text_bytes = items
        .iter()
        .fold(0usize, |total, item| total.saturating_add(item.len()));
    let bound = text_bytes
        .saturating_add(items.len().saturating_mul(128))
        .saturating_add(128)
        .max(1);
    popup_codec::DecodeOptions::new(
        bound,
        bound,
        bound,
        bound.saturating_mul(16),
        64,
        bound,
        items.len().max(1),
        text_bytes.max(1),
    )
}

fn bnc_options(source: &[u8]) -> PopUpMenuRewriteOptions {
    let bytes = source.len().max(1);
    PopUpMenuRewriteOptions {
        max_input_bytes: bytes,
        max_output_bytes: bytes.saturating_mul(4).max(1),
        max_fields: 128,
        max_work_bytes: bytes.saturating_mul(128).max(1),
        recursion_limit: 8,
        max_references: 64,
    }
}

fn map_bnc_error(error: PopUpMenuRewriteError) -> NativePopUpError {
    match error {
        PopUpMenuRewriteError::InvalidFormat(_) => NativePopUpError::InvalidSource,
        PopUpMenuRewriteError::InputBytes { .. }
        | PopUpMenuRewriteError::OutputBytes { .. }
        | PopUpMenuRewriteError::Fields { .. }
        | PopUpMenuRewriteError::Work { .. }
        | PopUpMenuRewriteError::Nesting { .. }
        | PopUpMenuRewriteError::References { .. }
        | PopUpMenuRewriteError::RetainedBytes { .. }
        | PopUpMenuRewriteError::ScratchBytes { .. } => NativePopUpError::Limit,
        PopUpMenuRewriteError::Allocations { .. } | PopUpMenuRewriteError::Allocation { .. } => {
            NativePopUpError::Allocation
        },
    }
}

pub(super) fn replace_message_preserving_header(
    archive: &mut Archive,
    object_identifier: u64,
    message_index: usize,
    payload: Vec<u8>,
    limits: Limits,
) -> Result<()> {
    let object = archive
        .object_mut(object_identifier)
        .ok_or(NativePopUpError::InvalidSource)?;
    let message_type = object
        .messages
        .get(message_index)
        .ok_or(NativePopUpError::InvalidSource)?
        .type_;
    object
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: message_type,
                data: payload,
            },
            limits,
        )
        .map_err(|_| NativePopUpError::Archive)?;
    Ok(())
}

/// Validate a message whose payload transition is not allowed to alter object
/// references.  The payload is still rewritten with the core raw-header
/// primitive, but stale aggregate or nested FieldInfo ownership must fail
/// before a private candidate can be published.
pub(super) fn validate_message_without_object_references(
    archive: &Archive,
    object_identifier: u64,
    message_index: usize,
) -> Result<()> {
    let object = unique_object(archive, object_identifier)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(NativePopUpError::InvalidSource)?;
    if !info.object_references.is_empty()
        || info.field_infos.iter().any(|field| {
            !field.object_references.is_empty()
                || field.effective_type() == FieldType::ObjectReference
        })
    {
        return Err(NativePopUpError::InvalidSource);
    }
    Ok(())
}

pub(super) fn replace_control_message_with_transition(
    archive: &mut Archive,
    object_identifier: u64,
    message_index: usize,
    source_payload: &[u8],
    payload: Vec<u8>,
    limits: Limits,
    preserve_aggregate_only: bool,
) -> Result<()> {
    let object = archive
        .object_mut(object_identifier)
        .ok_or(NativePopUpError::InvalidSource)?;
    let message_type = object
        .messages
        .get(message_index)
        .ok_or(NativePopUpError::InvalidSource)?
        .type_;
    let source_entries = decode_list_entries(source_payload, storage_options(source_payload))?;
    let after_entries = decode_list_entries(&payload, storage_options(&payload))?;
    let before = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(NativePopUpError::InvalidSource)?
        .object_references
        .clone();
    if duplicate_identifiers(&before) || before.contains(&0) {
        return Err(NativePopUpError::InvalidSource);
    }
    let before_entries = source_entries
        .entries
        .iter()
        .filter(|entry| entry.payload_kind == PayloadKind::ControlCellSpec)
        .map(|entry| {
            let payload = entry.payload.as_slice();
            let (spec, _) = control_codec::decode_any_cell_spec_with_report(
                payload,
                popup_options_for_bytes(payload),
            )
            .map_err(|_| NativePopUpError::Codec)?;
            if entry.ref_count == 0 {
                return Err(NativePopUpError::InvalidSource);
            }
            let identifier = match spec {
                control_codec::CellSpecSnapshot::Popup(spec) => {
                    let identifier = spec.popup_model().identifier();
                    if identifier == 0 {
                        return Err(NativePopUpError::InvalidSource);
                    }
                    Some(identifier)
                },
                control_codec::CellSpecSnapshot::Control(_) => None,
            };
            Ok((entry.key, identifier))
        })
        .collect::<Result<Vec<_>>>()?;
    let after_entries = after_entries
        .entries
        .iter()
        .filter(|entry| entry.payload_kind == PayloadKind::ControlCellSpec)
        .map(|entry| {
            let payload = entry.payload.as_slice();
            let (spec, _) = control_codec::decode_any_cell_spec_with_report(
                payload,
                popup_options_for_bytes(payload),
            )
            .map_err(|_| NativePopUpError::Codec)?;
            if entry.ref_count == 0 {
                return Err(NativePopUpError::InvalidSource);
            }
            let identifier = match spec {
                control_codec::CellSpecSnapshot::Popup(spec) => {
                    let identifier = spec.popup_model().identifier();
                    if identifier == 0 {
                        return Err(NativePopUpError::InvalidSource);
                    }
                    Some(identifier)
                },
                control_codec::CellSpecSnapshot::Control(_) => None,
            };
            Ok((entry.key, identifier))
        })
        .collect::<Result<Vec<_>>>()?;
    if duplicate_entry_keys(&before_entries) || duplicate_entry_keys(&after_entries) {
        return Err(NativePopUpError::InvalidSource);
    }
    let mut after = after_entries
        .iter()
        .filter_map(|(_, identifier)| *identifier)
        .collect::<Vec<_>>();
    if after.contains(&0) {
        return Err(NativePopUpError::InvalidSource);
    }
    // Multiple control-list entries may intentionally share one rooted popup
    // model while differing in initial selection. ArchiveInfo aggregates are
    // a set-like owner census; keep the per-entry duplicates in FieldInfo but
    // emit the model identifier only once in the aggregate transition.
    after.sort_unstable();
    after.dedup();
    let mut expected_before = before_entries
        .iter()
        .filter_map(|(_, identifier)| *identifier)
        .collect::<Vec<_>>();
    expected_before.sort_unstable();
    expected_before.dedup();
    let mut sorted_before = before.clone();
    sorted_before.sort_unstable();
    if sorted_before != expected_before {
        return Err(NativePopUpError::InvalidSource);
    }
    let info = object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or(NativePopUpError::InvalidSource)?;
    // Native Numbers producers commonly store only the aggregate popup
    // object reference for this list message.  An absent FieldInfo collection
    // is therefore an authoritative producer shape, not a missing proof.  It
    // must stay absent through a transition; when FieldInfo records exist we
    // validate every existing path and add a path only for a newly appended
    // popup entry.
    let aggregate_only = preserve_aggregate_only && info.field_infos.is_empty();
    // A builder-produced list may omit FieldInfo records only for entries
    // introduced by this transition. Existing rooted entries must already
    // have an exact [3,key] record; otherwise their aggregate and per-entry
    // ownership cannot be proved. Never invent a broadcast aggregate field.
    for (key, identifier) in &after_entries {
        // Scalar CellSpecs are owned by the BNC/list refcount and do not have
        // a PopupModel object-reference edge. Their list entries therefore
        // must not receive a synthetic ObjectReference FieldInfo.
        if identifier.is_none() {
            continue;
        }
        if !info
            .field_infos
            .iter()
            .any(|field| field.path.path.as_slice() == [3, *key])
        {
            if !aggregate_only
                && before_entries
                    .iter()
                    .any(|(before_key, _)| *before_key == *key)
            {
                return Err(NativePopUpError::InvalidSource);
            }
            if !aggregate_only {
                let mut field = FieldInfo::new(vec![3, *key]);
                field.r#type = Some(FieldType::ObjectReference);
                info.field_infos.push(field);
            }
        }
    }
    let mut field_changes = Vec::new();
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        let Some(path_key) = field
            .path
            .path
            .strip_prefix(&[3])
            .and_then(|rest| (rest.len() == 1).then_some(rest[0]))
        else {
            continue;
        };
        let old_id = before_entries
            .iter()
            .find(|(key, _)| *key == path_key)
            .and_then(|(_, id)| *id);
        let new_id = after_entries
            .iter()
            .find(|(key, _)| *key == path_key)
            .and_then(|(_, id)| *id);
        let entry_exists = before_entries
            .iter()
            .chain(after_entries.iter())
            .any(|(key, _)| *key == path_key);
        if old_id.is_none() && new_id.is_none() {
            if !entry_exists {
                return Err(NativePopUpError::InvalidSource);
            }
            if !field.object_references.is_empty()
                || field.effective_type() == FieldType::ObjectReference
            {
                return Err(NativePopUpError::InvalidSource);
            }
            continue;
        }
        let before_matches = match old_id {
            Some(identifier) => field.object_references.as_slice() == [identifier],
            None => field.object_references.is_empty(),
        };
        if field
            .r#type
            .is_some_and(|kind| kind != FieldType::ObjectReference)
            || !before_matches
        {
            return Err(NativePopUpError::InvalidSource);
        }
        field_changes.push(OwnedFieldChange {
            field_info_index,
            path: field.path.path.clone(),
            before: field.object_references.clone(),
            after: new_id.into_iter().collect(),
        });
    }
    if !aggregate_only {
        for (key, identifier) in &after_entries {
            if identifier.is_none() {
                continue;
            }
            if !field_changes
                .iter()
                .any(|change| change.path.as_slice() == [3, *key])
            {
                return Err(NativePopUpError::InvalidSource);
            }
        }
        for (key, identifier) in &before_entries {
            if identifier.is_none() {
                continue;
            }
            if !field_changes
                .iter()
                .any(|change| change.path.as_slice() == [3, *key])
            {
                return Err(NativePopUpError::InvalidSource);
            }
        }
    }
    if field_changes
        .iter()
        .map(|change| change.path.as_slice())
        .enumerate()
        .any(|(index, path)| {
            field_changes[index + 1..]
                .iter()
                .any(|other| other.path.as_slice() == path)
        })
    {
        return Err(NativePopUpError::InvalidSource);
    }
    let mut transitions = Vec::new();
    transitions
        .try_reserve_exact(field_changes.len())
        .map_err(|_| NativePopUpError::Allocation)?;
    for change in &field_changes {
        transitions.push(FieldObjectReferenceTransition {
            field_info_index: change.field_info_index,
            expected_path: change.path.as_slice(),
            before: change.before.as_slice(),
            after: change.after.as_slice(),
        });
    }
    let transition = ObjectReferenceTransition {
        aggregate_before: before.as_slice(),
        aggregate_after: after.as_slice(),
        fields: transitions.as_slice(),
    };
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: message_type,
                data: payload,
            },
            transition,
            limits,
        )
        .map_err(|_| NativePopUpError::Archive)?;
    // Remove control FieldInfo records for entries that no longer exist.  The
    // transition above has already validated and rewritten every retained
    // path, so this cleanup cannot shift an index used by the transition.
    if let Some(info) = object.archive_info.message_infos.get_mut(message_index) {
        info.field_infos.retain(|field| {
            let Some(key) = field
                .path
                .path
                .strip_prefix(&[3])
                .and_then(|rest| (rest.len() == 1).then_some(rest[0]))
            else {
                return true;
            };
            after_entries
                .iter()
                .any(|(after_key, identifier)| *after_key == key && identifier.is_some())
        });
    }
    let _ = source_payload;
    Ok(())
}

#[derive(Debug)]
struct OwnedFieldChange {
    field_info_index: usize,
    path: Vec<u32>,
    before: Vec<u64>,
    after: Vec<u64>,
}

fn popup_references_from_list(source: &[u8]) -> Result<Vec<u64>> {
    let list = decode_list_entries(source, storage_options(source))?;
    let mut references = Vec::new();
    for entry in list.entries {
        if entry.payload_kind != PayloadKind::ControlCellSpec {
            continue;
        }
        let (spec, _) = control_codec::decode_any_cell_spec_with_report(
            &entry.payload,
            popup_options_for_bytes(&entry.payload),
        )
        .map_err(|_| NativePopUpError::Codec)?;
        if let control_codec::CellSpecSnapshot::Popup(spec) = spec {
            let identifier = spec.popup_model().identifier();
            if identifier == 0 {
                return Err(NativePopUpError::InvalidSource);
            }
            if !references.contains(&identifier) {
                references.push(identifier);
            }
        }
    }
    Ok(references)
}

/// Return whether a known archive metadata edge or rooted control-list entry
/// points at a popup model.  The facade calls this over every package member
/// before allowing a cull; the native candidate still repeats the selected
/// archive check after its list rewrite.
pub(super) fn archive_has_popup_reference(archive: &Archive, identifier: u64) -> Result<bool> {
    for object in &archive.objects {
        for info in &object.archive_info.message_infos {
            if info.object_references.contains(&identifier)
                || info
                    .field_infos
                    .iter()
                    .any(|field| field.object_references.contains(&identifier))
            {
                return Ok(true);
            }
        }
        for message in &object.messages {
            if message.type_ != TABLE_DATA_LIST_TYPE {
                continue;
            }
            if table_data_list_type(&message.data, storage_options(&message.data))?
                != Some(LIST_CONTROL_CELL_SPEC)
            {
                continue;
            }
            let list = decode_list_entries(&message.data, storage_options(&message.data))?;
            for entry in list.entries {
                if entry.payload_kind != PayloadKind::ControlCellSpec {
                    continue;
                }
                let (spec, _) = control_codec::decode_any_cell_spec_with_report(
                    &entry.payload,
                    popup_options_for_bytes(&entry.payload),
                )
                .map_err(|_| NativePopUpError::Codec)?;
                if let control_codec::CellSpecSnapshot::Popup(spec) = spec
                    && spec.popup_model().identifier() == identifier
                {
                    return Ok(true);
                }
            }
        }
    }
    Ok(false)
}

/// Global-cull variant: an unknown ArchiveInfo/FieldInfo edge is itself an
/// ownership risk.  Callers use this only for non-selected package members so
/// a future opaque producer cannot keep a model alive while the selected native
/// archive is being rewritten.
pub(super) fn archive_has_popup_reference_strict(
    archive: &Archive,
    identifier: u64,
    limits: Limits,
) -> Result<bool> {
    let mut visitor = PopupReferenceVisitor {
        identifier,
        found: false,
    };
    for object in &archive.objects {
        object
            .inspect_references_with_policy_and_limits(
                &mut visitor,
                ArchiveReferencePolicy::RejectUnknownMetadata,
                limits,
            )
            .map_err(|_| NativePopUpError::Archive)?;
    }
    // ArchiveInfo does not carry the embedded CellSpec payloads of known
    // control-list objects.  Include that strict native census as well, so a
    // cross-component rooted list cannot be mistaken for an unowned orphan.
    Ok(visitor.found || archive_has_popup_reference(archive, identifier)?)
}

/// Verify object-level locality for a native rewrite that does not allocate or
/// cull archive objects. Every object outside `changed_identifiers` must keep
/// its complete source representation, including raw ArchiveInfo framing,
/// MessageInfo aggregate/FieldInfo records, unknown fields, and payload bytes.
/// The candidate may not silently add or remove an unlisted object.
pub(super) fn verify_archive_object_locality(
    source: &Archive,
    candidate: &Archive,
    changed_identifiers: &[u64],
) -> Result<()> {
    for source_object in &source.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(NativePopUpError::InvalidSource)?;
        if changed_identifiers.contains(&identifier) {
            continue;
        }
        let candidate_object = candidate
            .object(identifier)
            .ok_or(NativePopUpError::InvalidSource)?;
        if !source_object.same_content_ignoring_offsets(candidate_object) {
            return Err(NativePopUpError::InvalidSource);
        }
    }
    for candidate_object in &candidate.objects {
        let identifier = candidate_object
            .archive_info
            .identifier
            .ok_or(NativePopUpError::InvalidSource)?;
        if source.object(identifier).is_none() {
            return Err(NativePopUpError::InvalidSource);
        }
    }
    Ok(())
}

struct PopupReferenceVisitor {
    identifier: u64,
    found: bool,
}

impl ArchiveReferenceVisitor for PopupReferenceVisitor {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.kind == ArchiveReferenceKind::Object
            && occurrence.referenced_identifier == self.identifier
        {
            self.found = true;
        }
        Ok(())
    }
}

fn duplicate_identifiers(values: &[u64]) -> bool {
    values
        .iter()
        .enumerate()
        .any(|(index, value)| values[index + 1..].contains(value))
}

fn duplicate_entry_keys<T>(values: &[(u32, T)]) -> bool {
    values
        .iter()
        .enumerate()
        .any(|(index, (key, _))| values[index + 1..].iter().any(|(other, _)| other == key))
}

pub(super) fn tile_cell(source: &[u8], row: u32, column: u32) -> Result<&[u8]> {
    let view = WireView::parse(source).map_err(|_| NativePopUpError::InvalidSource)?;
    let mut selected = None;
    for field in view.fields() {
        if field.number() != 5 {
            continue;
        }
        let payload = field
            .canonical_payload()
            .map_err(|_| NativePopUpError::InvalidSource)?;
        let row_view = WireView::parse(payload).map_err(|_| NativePopUpError::InvalidSource)?;
        let row_index = row_view
            .fields()
            .find(|field| field.number() == 1)
            .and_then(|field| {
                decode_varint_from_bytes(field.payload())
                    .ok()
                    .map(|(value, _)| value)
            })
            .ok_or(NativePopUpError::InvalidSource)?;
        let row_index = u32::try_from(row_index).map_err(|_| NativePopUpError::InvalidSource)?;
        if row_index != row {
            continue;
        }
        if selected.is_some() {
            return Err(NativePopUpError::InvalidSource);
        }
        let cell_count = row_view
            .fields()
            .find(|field| field.number() == 2)
            .and_then(|field| {
                decode_varint_from_bytes(field.payload())
                    .ok()
                    .map(|(value, _)| value)
            })
            .ok_or(NativePopUpError::InvalidSource)?;
        let buffer = row_view
            .fields()
            .find(|field| field.number() == 6)
            .ok_or(NativePopUpError::InvalidSource)?
            .canonical_payload()
            .map_err(|_| NativePopUpError::InvalidSource)?;
        selected = Some(select_row_cell(&row_view, buffer, cell_count, column)?);
    }
    selected.ok_or(NativePopUpError::InvalidSource)
}

pub(super) fn patch_tile_cell(
    source: &[u8],
    row: u32,
    column: u32,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    let view = WireView::parse(source).map_err(|_| NativePopUpError::InvalidSource)?;
    let mut output = Vec::new();
    let mut selected = false;
    for field in view.fields() {
        if field.number() != 5 {
            output.extend_from_slice(field.raw());
            continue;
        }
        let payload = field
            .canonical_payload()
            .map_err(|_| NativePopUpError::InvalidSource)?;
        let row_view = WireView::parse(payload).map_err(|_| NativePopUpError::InvalidSource)?;
        let row_index = row_view
            .fields()
            .find(|candidate| candidate.number() == 1)
            .and_then(|candidate| {
                decode_varint_from_bytes(candidate.payload())
                    .ok()
                    .map(|(value, _)| value)
            })
            .and_then(|value| u32::try_from(value).ok())
            .ok_or(NativePopUpError::InvalidSource)?;
        if row_index != row {
            output.extend_from_slice(field.raw());
            continue;
        }
        if selected {
            return Err(NativePopUpError::InvalidSource);
        }
        let patched_row = patch_row_cell(payload, column, replacement)?;
        litchi_iwa_common::wire::append_length_delimited_field(&mut output, 5, &patched_row)
            .map_err(|_| NativePopUpError::InvalidSource)?;
        selected = true;
    }
    if !selected {
        return Err(NativePopUpError::InvalidSource);
    }
    Ok(output)
}

fn patch_row_cell(source: &[u8], column: u32, replacement: &[u8]) -> Result<Vec<u8>> {
    let view = WireView::parse(source).map_err(|_| NativePopUpError::InvalidSource)?;
    let cell_count = view
        .fields()
        .find(|field| field.number() == 2)
        .and_then(|field| {
            decode_varint_from_bytes(field.payload())
                .ok()
                .map(|(value, _)| value)
        })
        .ok_or(NativePopUpError::InvalidSource)?;
    let buffer = view
        .fields()
        .find(|field| field.number() == 6)
        .ok_or(NativePopUpError::InvalidSource)?
        .canonical_payload()
        .map_err(|_| NativePopUpError::InvalidSource)?;
    let (patched_buffer, patched_offsets) =
        patch_row_buffer(&view, buffer, cell_count, column, replacement)?;
    let mut output = Vec::new();
    let mut replaced = false;
    for field in view.fields() {
        if field.number() == 6 {
            if replaced {
                return Err(NativePopUpError::InvalidSource);
            }
            litchi_iwa_common::wire::append_length_delimited_field(&mut output, 6, &patched_buffer)
                .map_err(|_| NativePopUpError::InvalidSource)?;
            replaced = true;
        } else if field.number() == 7 {
            if let Some(offsets) = patched_offsets.as_deref() {
                litchi_iwa_common::wire::append_length_delimited_field(&mut output, 7, offsets)
                    .map_err(|_| NativePopUpError::InvalidSource)?;
            } else {
                output.extend_from_slice(field.raw());
            }
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !replaced {
        return Err(NativePopUpError::InvalidSource);
    }
    Ok(output)
}

/// Select one cell from a row's packed BNC storage.  Older files omit the
/// offset table only for a single-cell row; multi-cell rows must carry the
/// canonical 16-bit offset entries.  The wide-offset bit scales each entry in
/// four-byte units, matching the storage codec's census rules.
fn select_row_cell<'a>(
    row: &WireView<'a>,
    buffer: &'a [u8],
    cell_count: u64,
    column: u32,
) -> Result<&'a [u8]> {
    let count = usize::try_from(cell_count).map_err(|_| NativePopUpError::InvalidSource)?;
    let index = usize::try_from(column).map_err(|_| NativePopUpError::InvalidSource)?;
    if count == 0 {
        return Err(NativePopUpError::InvalidSource);
    }
    let offset_field = row.fields().find(|field| field.number() == 7);
    if count == 1 && offset_field.is_none() {
        if index != 0 {
            return Err(NativePopUpError::InvalidSource);
        }
        return Ok(buffer);
    }
    let offsets = offset_field
        .ok_or(NativePopUpError::UnsupportedDependency)?
        .canonical_payload()
        .map_err(|_| NativePopUpError::InvalidSource)?;
    let wide = row_wide_offsets(row)?;
    let unit = if wide { 4usize } else { 1usize };
    let slot_count = offsets.len() / 2;
    if slot_count < count || index >= slot_count {
        return Err(NativePopUpError::UnsupportedDependency);
    }
    let starts = decode_row_offsets(offsets, slot_count, unit)?;
    if starts.iter().flatten().count() != count {
        return Err(NativePopUpError::InvalidSource);
    }
    let start = starts[index].ok_or(NativePopUpError::UnsupportedDependency)?;
    let end = starts
        .iter()
        .skip(index + 1)
        .flatten()
        .next()
        .copied()
        .unwrap_or(buffer.len());
    if start > end || end > buffer.len() {
        return Err(NativePopUpError::InvalidSource);
    }
    Ok(&buffer[start..end])
}

fn decode_row_offsets(offsets: &[u8], count: usize, unit: usize) -> Result<Vec<Option<usize>>> {
    let expected = count
        .checked_mul(2)
        .ok_or(NativePopUpError::InvalidSource)?;
    if offsets.len() != expected {
        return Err(NativePopUpError::UnsupportedDependency);
    }
    let mut starts = Vec::with_capacity(count);
    let mut previous = None;
    for encoded in offsets.chunks_exact(2) {
        let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
        if raw == u16::MAX {
            starts.push(None);
            continue;
        }
        let start = usize::from(raw)
            .checked_mul(unit)
            .ok_or(NativePopUpError::InvalidSource)?;
        // Minimal/automatic cells have an empty BNC payload. Adjacent empty
        // cells therefore legitimately share the same start offset; only a
        // decreasing offset would make the packed row ambiguous.
        if previous.is_some_and(|prior| prior > start) {
            return Err(NativePopUpError::InvalidSource);
        }
        starts.push(Some(start));
        previous = Some(start);
    }
    Ok(starts)
}

fn row_wide_offsets(row: &WireView<'_>) -> Result<bool> {
    let Some(field) = row.fields().find(|field| field.number() == 8) else {
        return Ok(false);
    };
    let (value, consumed) =
        decode_varint_from_bytes(field.payload()).map_err(|_| NativePopUpError::InvalidSource)?;
    if consumed != field.payload().len() || value > 1 {
        return Err(NativePopUpError::InvalidSource);
    }
    Ok(value != 0)
}

fn patch_row_buffer(
    row: &WireView<'_>,
    buffer: &[u8],
    cell_count: u64,
    column: u32,
    replacement: &[u8],
) -> Result<(Vec<u8>, Option<Vec<u8>>)> {
    let count = usize::try_from(cell_count).map_err(|_| NativePopUpError::InvalidSource)?;
    let index = usize::try_from(column).map_err(|_| NativePopUpError::InvalidSource)?;
    if count == 0 {
        return Err(NativePopUpError::InvalidSource);
    }
    let offset_field = row.fields().find(|field| field.number() == 7);
    if count == 1 && offset_field.is_none() {
        if index != 0 {
            return Err(NativePopUpError::InvalidSource);
        }
        return Ok((replacement.to_owned(), None));
    }
    let offsets = offset_field
        .ok_or(NativePopUpError::UnsupportedDependency)?
        .canonical_payload()
        .map_err(|_| NativePopUpError::InvalidSource)?;
    let wide = row_wide_offsets(row)?;
    let unit = if wide { 4usize } else { 1usize };
    let slot_count = offsets.len() / 2;
    if slot_count < count || index >= slot_count {
        return Err(NativePopUpError::UnsupportedDependency);
    }
    let starts = decode_row_offsets(offsets, slot_count, unit)?;
    if starts.iter().flatten().count() != count {
        return Err(NativePopUpError::InvalidSource);
    }
    let start = starts[index].ok_or(NativePopUpError::UnsupportedDependency)?;
    let end = starts
        .iter()
        .skip(index + 1)
        .flatten()
        .next()
        .copied()
        .unwrap_or(buffer.len());
    if start > end || end > buffer.len() {
        return Err(NativePopUpError::InvalidSource);
    }
    let old_len = end - start;
    let delta = replacement.len() as isize - old_len as isize;
    let capacity = if delta.is_negative() {
        buffer
            .len()
            .checked_sub(delta.unsigned_abs())
            .ok_or(NativePopUpError::InvalidSource)?
    } else {
        buffer
            .len()
            .checked_add(delta as usize)
            .ok_or(NativePopUpError::InvalidSource)?
    };
    let mut output = Vec::with_capacity(capacity);
    output.extend_from_slice(&buffer[..start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&buffer[end..]);
    let mut replacement_offsets = None;
    if delta != 0 {
        // Keep the offset table source-preserving while updating only the
        // canonical starts following the selected cell.  Re-encode offsets
        // with the same width/unit and reject values that no longer fit.
        let mut encoded_offsets = Vec::with_capacity(offsets.len());
        for (offset_index, encoded) in offsets.chunks_exact(2).enumerate() {
            let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
            if raw == u16::MAX {
                encoded_offsets.extend_from_slice(&u16::MAX.to_le_bytes());
                continue;
            }
            let original = usize::from(raw)
                .checked_mul(unit)
                .ok_or(NativePopUpError::InvalidSource)?;
            let adjusted = if offset_index > index {
                if delta.is_negative() {
                    original.checked_sub(delta.unsigned_abs())
                } else {
                    original.checked_add(delta as usize)
                }
            } else {
                Some(original)
            }
            .ok_or(NativePopUpError::InvalidSource)?;
            if !adjusted.is_multiple_of(unit) || adjusted / unit > usize::from(u16::MAX) {
                return Err(NativePopUpError::UnsupportedDependency);
            }
            encoded_offsets.extend_from_slice(
                &u16::try_from(adjusted / unit)
                    .map_err(|_| NativePopUpError::UnsupportedDependency)?
                    .to_le_bytes(),
            );
        }
        replacement_offsets = Some(encoded_offsets);
    }
    Ok((output, replacement_offsets))
}
