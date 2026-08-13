//! Exact-source focused Keynote soundtrack-settings transactions.

mod media;
mod rewrite;

use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::ExactArtifacts};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::ArchiveObject;
use litchi_iwa_protos::{
    keynote_document_codec, keynote_soundtrack_settings_codec as soundtrack_codec,
};

use super::{DOCUMENT_MESSAGE_TYPE, Package, ReadError, SHOW_MESSAGE_TYPE};
use crate::soundtrack::{Commit, Diagnostics, Edit, Error, LimitKind, Mode, Patch, Settings};
use media::validate_media_closure;
use rewrite::{rewrite_and_verify, validate_component_framing, verify_candidate};

const SOUNDTRACK_MESSAGE_TYPE: u32 = 21;
const DOCUMENT_SHOW_FIELD: u32 = 2;
const SHOW_SOUNDTRACK_FIELD: u32 = 17;
const SOUNDTRACK_MEDIA_FIELD: u32 = 3;
const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const PACKAGE_METADATA_TYPE: u32 = 11_006;

pub(crate) struct Prepared<'a> {
    selection: Selection<'a>,
    budget: TransactionBudget,
}

struct Selection<'a> {
    root_component: &'a str,
    show_component: &'a str,
    soundtrack_component: &'a str,
    show_identifier: u64,
    soundtrack_identifier: u64,
    root: &'a ArchiveObject,
    root_message_index: usize,
    show: &'a ArchiveObject,
    show_message_index: usize,
    soundtrack: &'a ArchiveObject,
    soundtrack_message_index: usize,
    soundtrack_payload: &'a [u8],
    settings: Settings,
}

struct MediaClosureState<'a> {
    payload_occurrences: usize,
    component_declarations: usize,
    owner_occurrences: usize,
    owner_count: usize,
    data_declarations: usize,
    filename: Option<&'a [u8]>,
    materialized_length: Option<usize>,
}

#[derive(Clone, Copy)]
struct RawField<'a> {
    number: u32,
    wire: u8,
    varint: Option<u64>,
    bytes: Option<&'a [u8]>,
    raw: &'a [u8],
}

#[derive(Clone, Copy)]
struct ReferenceFacts {
    identifier: u64,
    external: Option<bool>,
    fields: usize,
}

#[derive(Clone, Copy)]
struct ReopenCost {
    work: usize,
    references: usize,
}

#[derive(Clone, Copy)]
struct TransactionBudget {
    fields: usize,
    work: usize,
    references: usize,
    max_fields: usize,
    max_work: usize,
    max_references: usize,
    max_nesting: usize,
}

impl TransactionBudget {
    fn new(package: &Package) -> Result<Self, Error> {
        let limits = package.wire_limits().map_err(map_wire_error)?;
        Ok(Self {
            fields: 0,
            work: 0,
            references: 0,
            max_fields: limits.max_fields(),
            max_work: limits.max_rewrite_work(),
            max_references: package.semantic_limits().max_references(),
            max_nesting: limits.max_nesting(),
        })
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), Error> {
        self.fields = checked_charge(self.fields, amount, self.max_fields, LimitKind::WireFields)?;
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), Error> {
        self.work = checked_charge(self.work, amount, self.max_work, LimitKind::WireWork)?;
        Ok(())
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), Error> {
        self.references = checked_charge(
            self.references,
            amount,
            self.max_references,
            LimitKind::References,
        )?;
        Ok(())
    }

    fn remaining_fields(self) -> usize {
        self.max_fields.saturating_sub(self.fields)
    }

    fn remaining_work(self) -> usize {
        self.max_work.saturating_sub(self.work)
    }

    fn merge_codec(&mut self, report: soundtrack_codec::DecodeReport) -> Result<(), Error> {
        self.charge_fields(report.fields())?;
        self.charge_work(report.work_bytes())
    }

    fn require_depth(self, depth: usize) -> Result<(), Error> {
        if depth > self.max_nesting {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }
}

impl Package {
    /// Read the presentation soundtrack's playback settings.
    ///
    /// `None` means the presentation has no existing soundtrack settings.
    /// `Some(Settings::default())` means settings exist while both optional
    /// playback values are absent. These cases are intentionally distinct.
    ///
    /// Soundtrack media is orthogonal and is neither exposed nor changed by
    /// this focused value.
    ///
    /// # Costs
    ///
    /// Performs a bounded focused read of presentation and soundtrack settings
    /// metadata. It does not initialize the full slide graph, publish output,
    /// or copy soundtrack media.
    ///
    /// # Errors
    ///
    /// Returns a typed source, resource, or allocation error when the
    /// soundtrack settings cannot be read safely. An absent soundtrack is
    /// returned as `Ok(None)`, not as an error.
    pub fn soundtrack_settings(&self) -> Result<Option<Settings>, Error> {
        read_settings(self)
    }

    /// Start an exact immutable edit of existing soundtrack playback settings.
    ///
    /// This transaction can update only volume and mode for an existing
    /// soundtrack. It neither creates nor deletes soundtrack settings or media
    /// resources, and it does not expose media entries.
    ///
    /// # Costs
    ///
    /// Performs the same focused read as [`Self::soundtrack_settings`] and
    /// retains compact semantic settings plus an immutable package borrow.
    /// [`Edit::set`] is constant-time and allocation-free.
    ///
    /// # Errors
    ///
    /// Returns [`Error::SoundtrackNotFound`] when no soundtrack settings exist,
    /// or another typed source, resource, or allocation error.
    pub fn edit_soundtrack_settings(&self) -> Result<Edit<'_>, Error> {
        let mut budget = TransactionBudget::new(self)?;
        let selection = select(self, &mut budget)?.ok_or(Error::SoundtrackNotFound)?;
        let before = selection.settings;
        Ok(Edit {
            source: self,
            before,
            settings: before,
            prepared: Prepared { selection, budget },
        })
    }

    /// Apply an exact-source-checked soundtrack-settings patch.
    ///
    /// # Costs
    ///
    /// Exact authorization uses an allocation-identity fast path, otherwise
    /// one complete package comparison. A no-op returns before focused
    /// selection or reopening; a change reopens and verifies the retained
    /// target once. Soundtrack media and rendering previews remain outside the
    /// changed settings scope.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for a different exact source package
    /// snapshot or mismatched prior semantic state, or a typed source, resource,
    /// allocation, or verification error for a rejected retained target.
    pub fn apply_soundtrack_settings(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics {
                    changed: false,
                    touched_components: 0,
                    full_reparse_performed: false,
                },
            });
        }
        let mut budget = TransactionBudget::new(self)?;
        let source_selection = select(self, &mut budget)?.ok_or(Error::PatchConflict)?;
        if source_selection.settings != patch.before {
            return Err(Error::PatchConflict);
        }
        validate_mutation_source(self, &source_selection, &mut budget)?;
        let target_bytes = patch.artifacts.target();
        budget.charge_work(patch.target_reopen_work)?;
        budget.charge_references(patch.target_reopen_references)?;
        let candidate = Package::from_source_with_options(target_bytes, self.state.options)
            .map_err(map_read_error)?;
        let touched = verify_candidate(
            self,
            &candidate,
            &source_selection,
            patch.after,
            &mut budget,
        )?;
        if touched != patch.touched_components {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                touched_components: touched,
                full_reparse_performed: true,
            },
        })
    }
}

pub(crate) fn commit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = physical_source(edit.source)?;
    if edit.before == edit.settings {
        return Ok(Commit {
            package: edit.source.snapshot(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&source), source),
                before: edit.before,
                after: edit.settings,
                touched_components: 0,
                source_reopen_work: 0,
                target_reopen_work: 0,
                source_reopen_references: 0,
                target_reopen_references: 0,
            },
            diagnostics: Diagnostics {
                changed: false,
                touched_components: 0,
                full_reparse_performed: false,
            },
        });
    }
    let Prepared {
        selection,
        mut budget,
    } = edit.prepared;
    validate_mutation_source(edit.source, &selection, &mut budget)?;
    rewrite_and_verify(
        edit.source,
        source,
        &selection,
        edit.before,
        edit.settings,
        &mut budget,
    )
}

fn validate_mutation_source(
    package: &Package,
    selection: &Selection<'_>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let catalog = physical_catalog(package)?;
    budget.charge_work(catalog.components().len())?;
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    for (object, index) in [
        (selection.root, selection.root_message_index),
        (selection.show, selection.show_message_index),
        (selection.soundtrack, selection.soundtrack_message_index),
    ] {
        charge_message_info(object, index, budget)?;
        validate_selected_metadata(object, index)?;
    }
    validate_object_reference_metadata(
        selection.root,
        selection.root_message_index,
        selection.show_identifier,
        &[DOCUMENT_SHOW_FIELD],
    )?;
    validate_object_reference_metadata(
        selection.show,
        selection.show_message_index,
        selection.soundtrack_identifier,
        &[SHOW_SOUNDTRACK_FIELD],
    )?;
    validate_soundtrack_metadata(selection)?;
    validate_media_closure(package, selection, budget)?;
    let wire_limits = package.wire_limits().map_err(map_wire_error)?;
    validate_reference_role_disjointness(selection, wire_limits, budget)?;
    let names = [
        selection.root_component,
        selection.show_component,
        selection.soundtrack_component,
    ];
    for (index, name) in names.into_iter().enumerate() {
        if names[..index].contains(&name) {
            continue;
        }
        if name == selection.soundtrack_component {
            continue;
        }
        validate_component_framing(package, name, budget)?;
    }
    Ok(())
}

fn validate_selected_metadata(object: &ArchiveObject, index: usize) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_object_reference_metadata(
    object: &ArchiveObject,
    index: usize,
    identifier: u64,
    path: &[u32],
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Error::InvalidSource)?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource);
    }
    let mut selected_path = false;
    for field in &info.field_infos {
        if field.path.as_slice() == path {
            if selected_path
                || field.r#type != Some(litchi_iwa_core::FieldType::ObjectReference)
                || field.object_references.as_slice() != [identifier]
            {
                return Err(Error::InvalidSource);
            }
            selected_path = true;
        } else if field.object_references.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn validate_soundtrack_metadata(selection: &Selection<'_>) -> Result<(), Error> {
    let info = selection
        .soundtrack
        .archive_info
        .message_infos
        .get(selection.soundtrack_message_index)
        .ok_or(Error::InvalidSource)?;
    if !info.object_references.is_empty() {
        return Err(Error::InvalidSource);
    }
    let mut media_path = false;
    for field in &info.field_infos {
        if !field.object_references.is_empty() {
            return Err(Error::InvalidSource);
        }
        if field.path.as_slice() == [SOUNDTRACK_MEDIA_FIELD] {
            if media_path
                || field.r#type != Some(litchi_iwa_core::FieldType::DataReference)
                || field.data_references != info.data_references
            {
                return Err(Error::InvalidSource);
            }
            media_path = true;
        } else if !field.data_references.is_empty() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn charge_message_info(
    object: &ArchiveObject,
    index: usize,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Error::InvalidSource)?;
    let aggregate_references = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(
        1usize
            .checked_add(aggregate_references)
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.charge_references(aggregate_references)?;
    for field in &info.field_infos {
        budget.charge_work(
            1usize
                .checked_add(field.path.path.len())
                .and_then(|amount| amount.checked_add(field.object_references.len()))
                .and_then(|amount| amount.checked_add(field.data_references.len()))
                .ok_or(Error::InvalidSource)?,
        )?;
        budget.charge_references(
            field
                .object_references
                .len()
                .checked_add(field.data_references.len())
                .ok_or(Error::InvalidSource)?,
        )?;
    }
    Ok(())
}

fn validate_reference_role_disjointness(
    selection: &Selection<'_>,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let (_, root_payload) = selected_message(selection.root, DOCUMENT_MESSAGE_TYPE)?;
    reject_selected_identifier_in_reference_fields(
        root_payload,
        &[4],
        selection.soundtrack_identifier,
        limits,
        budget,
    )?;
    let (_, show_payload) = selected_message(selection.show, SHOW_MESSAGE_TYPE)?;
    reject_selected_identifier_in_reference_fields(
        show_payload,
        &[1, 2, 5, 7, 19],
        selection.soundtrack_identifier,
        limits,
        budget,
    )?;
    budget.charge_work(show_payload.len())?;
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let show = WireView::parse_with_limits(show_payload, bounded_limits).map_err(map_wire_error)?;
    budget.charge_fields(show.len())?;
    let mut slide_tree_payload = None;
    for field in show.fields() {
        if field.number() != 3 {
            continue;
        }
        if slide_tree_payload.is_some() || field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        slide_tree_payload = Some(field.payload());
    }
    let tree_payload = slide_tree_payload.ok_or(Error::InvalidSource)?;
    reject_selected_identifier_in_reference_fields(
        tree_payload,
        &[1, 2],
        selection.soundtrack_identifier,
        limits,
        budget,
    )
}

fn reject_selected_identifier_in_reference_fields(
    source: &[u8],
    reference_fields: &[u32],
    selected: u64,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    budget.charge_work(source.len())?;
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let view = WireView::parse_with_limits(source, bounded_limits).map_err(map_wire_error)?;
    let outer_fields = view.len();
    budget.charge_fields(outer_fields)?;
    for field in view.fields() {
        if reference_fields.contains(&field.number()) {
            if field.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
            field.validate_canonical_framing().map_err(map_wire_error)?;
            let reference = strict_reference(field.payload(), limits, budget)?;
            if reference.identifier == selected || reference.external == Some(true) {
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok(())
}

fn read_settings(package: &Package) -> Result<Option<Settings>, Error> {
    let mut budget = TransactionBudget::new(package)?;
    Ok(select(package, &mut budget)?.map(|selection| selection.settings))
}

fn select<'a>(
    package: &'a Package,
    budget: &mut TransactionBudget,
) -> Result<Option<Selection<'a>>, Error> {
    let catalog = physical_catalog(package)?;
    let mut roots = catalog
        .components()
        .iter()
        .filter(|component| component.name().rsplit('/').next() == Some("Document.iwa"));
    let root_component = roots.next().ok_or(Error::InvalidSource)?;
    if roots.next().is_some() {
        return Err(Error::InvalidSource);
    }
    let root = root_component
        .archive()
        .object(1)
        .ok_or(Error::InvalidSource)?;
    let (root_message_index, root_payload) = selected_message(root, DOCUMENT_MESSAGE_TYPE)?;
    let wire_limits = package.wire_limits().map_err(map_wire_error)?;
    let nesting =
        u32::try_from(wire_limits.max_nesting()).map_err(|_error| Error::InvalidSource)?;
    let strict_root_reference =
        strict_optional_reference(root_payload, DOCUMENT_SHOW_FIELD, wire_limits, budget)?
            .ok_or(Error::InvalidSource)?;
    budget.charge_fields(strict_root_reference.fields.saturating_mul(2))?;
    budget.charge_work(root_payload.len().saturating_mul(4))?;
    let root_reference = keynote_document_codec::decode_show_reference(
        root_payload,
        keynote_document_codec::DecodeOptions::new(root_payload.len(), nesting)
            .with_max_fields(wire_limits.max_fields())
            .with_max_work_bytes(wire_limits.max_rewrite_work()),
    )
    .map_err(|error| map_document_codec_error(error, budget))?;
    let show_identifier = root_reference.identifier();
    if show_identifier != strict_root_reference.identifier
        || root_reference.deprecated_is_external() != strict_root_reference.external
        || show_identifier == 0
        || root_reference.deprecated_is_external() == Some(true)
    {
        return Err(Error::InvalidSource);
    }
    let (show_component_name, show) = unique_object(package, show_identifier)?;
    let (show_message_index, show_payload) = selected_message(show, SHOW_MESSAGE_TYPE)?;
    let show_reference =
        strict_optional_reference(show_payload, SHOW_SOUNDTRACK_FIELD, wire_limits, budget)?;
    let Some(soundtrack_reference) = show_reference else {
        return Ok(None);
    };
    let soundtrack_identifier = soundtrack_reference.identifier;
    if soundtrack_identifier == 0
        || soundtrack_reference.external == Some(true)
        || soundtrack_identifier == show_identifier
        || soundtrack_identifier == 1
    {
        return Err(Error::InvalidSource);
    }
    let (soundtrack_component_name, soundtrack) = unique_object(package, soundtrack_identifier)?;
    let (soundtrack_message_index, soundtrack_payload) =
        selected_message(soundtrack, SOUNDTRACK_MESSAGE_TYPE)?;
    let codec_options = soundtrack_codec::DecodeOptions::new(
        soundtrack_payload.len(),
        budget.remaining_fields(),
        budget.remaining_work(),
        nesting,
    );
    let info = soundtrack
        .archive_info
        .message_infos
        .get(soundtrack_message_index)
        .ok_or(Error::InvalidSource)?;
    budget.charge_references(info.data_references.len())?;
    let mut media_index = 0usize;
    let mut media_match = true;
    let (snapshot, report) = soundtrack_codec::decode_soundtrack_settings_with_media_report(
        soundtrack_payload,
        codec_options,
        &mut |identifier| {
            if info.data_references.get(media_index) != Some(&identifier) {
                media_match = false;
            }
            media_index = media_index.saturating_add(1);
            Ok(())
        },
    )
    .map_err(|error| map_soundtrack_codec_error(error, budget))?;
    if !media_match
        || media_index != info.data_references.len()
        || report.media_references() != media_index
    {
        return Err(Error::InvalidSource);
    }
    budget.merge_codec(report)?;
    let settings = Settings::new(snapshot.volume(), snapshot.mode_raw().map(Mode::from_raw))
        .map_err(|_error| Error::InvalidSource)?;
    Ok(Some(Selection {
        root_component: root_component.name(),
        show_component: show_component_name,
        soundtrack_component: soundtrack_component_name,
        show_identifier,
        soundtrack_identifier,
        root,
        root_message_index,
        show,
        show_message_index,
        soundtrack,
        soundtrack_message_index,
        soundtrack_payload,
        settings,
    }))
}

fn strict_optional_reference(
    source: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<Option<ReferenceFacts>, Error> {
    budget.charge_work(source.len())?;
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let view = WireView::parse_with_limits(source, bounded_limits).map_err(map_wire_error)?;
    let outer_fields = view.len();
    budget.charge_fields(outer_fields)?;
    let mut selected = None;
    for field in view.fields() {
        if field.number() != field_number {
            continue;
        }
        if selected.is_some() || field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let mut facts = strict_reference(field.payload(), limits, budget)?;
        facts.fields = facts
            .fields
            .checked_add(outer_fields)
            .ok_or(Error::InvalidSource)?;
        selected = Some(facts);
    }
    Ok(selected)
}

fn strict_reference(
    source: &[u8],
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<ReferenceFacts, Error> {
    budget.charge_work(source.len())?;
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let view = WireView::parse_with_limits(source, bounded_limits).map_err(map_wire_error)?;
    budget.charge_fields(view.len())?;
    let field_count = view.len();
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    for field in view.fields() {
        field.validate_canonical_key().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if identifier.is_some() || field.wire_type() != 0 {
                    return Err(Error::InvalidSource);
                }
                identifier = Some(canonical_varint(field.payload())?);
            },
            2 => {
                if deprecated_type.is_some() || field.wire_type() != 0 {
                    return Err(Error::InvalidSource);
                }
                let raw = canonical_varint(field.payload())?;
                if raw > i32::MAX.cast_unsigned().into() && raw < 0xffff_ffff_8000_0000 {
                    return Err(Error::InvalidSource);
                }
                deprecated_type = Some(raw);
            },
            3 => {
                if external.is_some() || field.wire_type() != 0 {
                    return Err(Error::InvalidSource);
                }
                external = Some(match canonical_varint(field.payload())? {
                    0 => false,
                    1 => true,
                    _ => return Err(Error::InvalidSource),
                });
            },
            _ => {},
        }
    }
    let selected_identifier = identifier.ok_or(Error::InvalidSource)?;
    budget.charge_references(1)?;
    Ok(ReferenceFacts {
        identifier: selected_identifier,
        external,
        fields: field_count,
    })
}

fn canonical_varint(source: &[u8]) -> Result<u64, Error> {
    let (value, consumed) =
        decode_varint_from_bytes(source).map_err(|_error| Error::InvalidSource)?;
    if consumed != source.len() || consumed != encoded_len(value) {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

fn selected_message(object: &ArchiveObject, kind: u32) -> Result<(usize, &[u8]), Error> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(Error::InvalidSource);
    }
    let mut selected = None;
    for (index, (message, info)) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
        .enumerate()
    {
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(Error::InvalidSource);
        }
        if message.type_ == kind && selected.replace((index, message.data.as_slice())).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

fn unique_object(package: &Package, identifier: u64) -> Result<(&str, &ArchiveObject), Error> {
    package
        .object_with_component(identifier)
        .ok_or(Error::InvalidSource)
}

fn remaining_wire_limits(
    limits: WireLimits,
    budget: &TransactionBudget,
) -> Result<WireLimits, Error> {
    if budget.remaining_fields() == 0 {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: budget.max_fields.saturating_add(1) as u64,
            maximum: budget.max_fields as u64,
        });
    }
    if budget.remaining_work() == 0 {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: budget.max_work.saturating_add(1) as u64,
            maximum: budget.max_work as u64,
        });
    }
    limits
        .with_fields(budget.remaining_fields())
        .and_then(|bounded| bounded.with_rewrite_work(budget.remaining_work()))
        .map_err(map_wire_error)
}

fn checked_charge(
    current: usize,
    amount: usize,
    maximum: usize,
    kind: LimitKind,
) -> Result<usize, Error> {
    let observed = current.saturating_add(amount);
    if observed > maximum {
        return Err(Error::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        });
    }
    Ok(observed)
}

#[allow(
    clippy::unnecessary_wraps,
    reason = "the internal-source feature adds a semantic-source error path"
)]
fn physical_source(package: &Package) -> Result<Arc<[u8]>, Error> {
    let source = match &package.state.source {
        super::PhysicalSource::Package(source) => source,
        super::PhysicalSource::Semantic(_) => return Err(Error::UnsupportedSource),
    };
    Ok(source.shared_source())
}

#[allow(
    clippy::unnecessary_wraps,
    reason = "the internal-source feature adds a semantic-source error path"
)]
fn physical_catalog(package: &Package) -> Result<&SourceCatalog, Error> {
    match &package.state.source {
        super::PhysicalSource::Package(source) => Ok(source),
        super::PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_document_codec_error(
    error: keynote_document_codec::DecodeError,
    budget: &TransactionBudget,
) -> Error {
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            keynote_document_codec::WireResourceLimit::Bytes { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::WireBytes,
                    observed: observed as u64,
                    maximum: maximum as u64,
                }
            },
            keynote_document_codec::WireResourceLimit::Nesting { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => Error::InvalidSource,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: budget.fields.saturating_add(observed) as u64,
            maximum: budget.fields.saturating_add(maximum) as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: budget.work.saturating_add(observed) as u64,
            maximum: budget.work.saturating_add(maximum) as u64,
        };
    }
    Error::InvalidSource
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_soundtrack_codec_error(
    error: soundtrack_codec::DecodeError,
    budget: &TransactionBudget,
) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource;
    };
    match limit {
        soundtrack_codec::DecodeLimit::Bytes { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        soundtrack_codec::DecodeLimit::Fields { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: budget.fields.saturating_add(observed) as u64,
            maximum: budget.fields.saturating_add(maximum) as u64,
        },
        soundtrack_codec::DecodeLimit::Work { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: budget.work.saturating_add(observed) as u64,
            maximum: budget.work.saturating_add(maximum) as u64,
        },
        soundtrack_codec::DecodeLimit::Nesting { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::Entries,
                super::SemanticLimitKind::Slides => LimitKind::Slides,
                super::SemanticLimitKind::References => LimitKind::References,
                super::SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                super::SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                super::SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(archive_error) => map_archive_error(archive_error),
        ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::NotKeynote
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_) => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::EntryBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => Error::InvalidSource,
    }
}
