//! Exact-source aggregate transactions for settings stored on a Pages section.

use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{SourceCatalog, package::Entry};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{
        NestedFieldEdit, NestedFieldReplacement, WireView, patch_nested_fields_batched_with_limits,
    },
};
use litchi_iwa_core::SnappyStream;
use litchi_iwa_protos::pages_section_codec::{
    DecodeLimit, DecodeOptions, DecodeReport, SectionSettingsSnapshot,
    decode_section_settings_with_report,
};

use super::{Package, section_transaction as transaction};
use crate::{
    SectionSelector,
    section::{
        PageNumber, PageNumbering, Settings, Start,
        settings::{Commit, Diagnostics, Edit, Error, Patch, Path},
    },
};

const INHERIT_FIELD: u32 = 17;
const FIRST_PAGE_DIFFERENT_FIELD: u32 = 18;
const EVEN_ODD_DIFFERENT_FIELD: u32 = 19;
const START_FIELD: u32 = 20;
const NUMBERING_FIELD: u32 = 21;
const STARTING_PAGE_FIELD: u32 = 22;
const FIRST_TEMPLATE_FIELD: u32 = 23;
const EVEN_TEMPLATE_FIELD: u32 = 24;
const ODD_TEMPLATE_FIELD: u32 = 25;
const NAME_FIELD: u32 = 26;
const HIDE_FIRST_FIELD: u32 = 28;

impl Package {
    /// Read every presence-preserving setting stored directly on one section.
    pub fn section_settings<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Settings, Error> {
        let position = transaction::resolve_position(self, selector)?;
        settings_at_position(self, position)
    }

    /// Begin one immutable aggregate section-settings transaction.
    pub fn edit_section_settings<'package, 'selector>(
        &'package self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Edit<'package>, Error> {
        let position = transaction::resolve_position(self, selector)?;
        let before = settings_at_position(self, position)?;
        Ok(Edit::from_package_parts(self, position, before))
    }

    /// Apply a reversible aggregate patch to this exact package artifact.
    pub fn apply_section_settings(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}

pub(crate) fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let (source, position, before, after) = edit.into_package_parts();
    after.validate().map_err(Error::InvalidSettings)?;
    let source_artifact = source.state.source.shared_source();
    let source_fingerprint = transaction::fingerprint(source_artifact.as_ref());
    if before == after {
        let patch = Patch::from_package_parts(
            Arc::clone(&source_artifact),
            source_artifact,
            source_fingerprint,
            source_fingerprint,
            position,
            before,
            after,
            None,
            None,
            0,
            0,
            0,
        );
        return Ok(Commit::from_parts(
            source.snapshot(),
            patch,
            Diagnostics::unchanged(),
        ));
    }
    if settings_at_position(source, position)? != before {
        return Err(Error::InvalidSource {
            path: Path::section(position),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(Error::UnsupportedSource {
            path: Path::section(position),
        });
    }

    let mut budget = transaction::TransactionBudget::new(source)?;
    let target = transaction::resolve_target(source, position, &mut budget)?;
    let payload = transaction::selected_payload(source, target)?;
    let decoded = decode_settings(source, payload, Path::section(position), &mut budget)?;
    if decoded != before {
        return Err(Error::InvalidSource {
            path: Path::section(position),
        });
    }
    validate_changed_dependencies(source, target, &before, &after, &mut budget)?;
    let rewritten = rewrite_payload(source, payload, &before, &after, &mut budget)?;
    let (candidate, stats) =
        transaction::rewrite_package(source, target, rewritten, false, &mut budget)?;
    if stats.touched_components != 1
        || stats.deleted_previews != 0
        || stats.source_layout_state.is_some()
        || stats.target_layout_state.is_some()
        || stats.source_preview_count != stats.target_preview_count
    {
        return Err(Error::Verification {
            path: Path::section(position),
        });
    }
    let (_candidate_target, candidate_settings) =
        settings_at_position_with_budget(&candidate, position, &mut budget)?;
    if candidate_settings != after {
        return Err(Error::Verification {
            path: Path::section(position),
        });
    }
    verify_unselected_members(source, &candidate, target, &before, &after)?;
    budget.settle_transaction_reservation();

    let target_artifact = candidate.state.source.shared_source();
    let target_fingerprint = transaction::fingerprint(target_artifact.as_ref());
    let patch = Patch::from_package_parts(
        source_artifact,
        target_artifact,
        source_fingerprint,
        target_fingerprint,
        position,
        before,
        after,
        None,
        None,
        stats.source_preview_count,
        stats.target_preview_count,
        stats.touched_components,
    );
    Ok(Commit::from_parts(
        candidate,
        patch,
        Diagnostics::published(stats.touched_components, 0, true),
    ))
}

fn apply_patch(source: &Package, patch: &Patch) -> Result<Commit, Error> {
    if source.source_bytes() != patch.source_artifact().as_ref() {
        return Err(Error::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(Commit::from_parts(
            source.snapshot(),
            patch.clone(),
            Diagnostics::unchanged(),
        ));
    }
    let mut budget = transaction::TransactionBudget::new(source)?;
    let (source_target, source_settings) =
        settings_at_position_with_budget(source, patch.position(), &mut budget)?;
    if source_settings != *patch.before() {
        return Err(Error::PatchConflict);
    }
    validate_changed_dependencies(
        source,
        source_target,
        patch.before(),
        patch.after(),
        &mut budget,
    )?;

    budget.charge_transaction_work(
        patch
            .target_artifact()
            .len()
            .saturating_mul(4)
            .saturating_add(128),
        Path::Package,
    )?;
    let catalog = SourceCatalog::from_shared_bytes_with_limits(
        Arc::clone(patch.target_artifact()),
        source.state.source.limits(),
    )
    .map_err(transaction::map_archive_error)?;
    let candidate =
        Package::from_source_catalog(catalog).map_err(transaction::map_package_error)?;
    let (_candidate_target, candidate_settings) =
        settings_at_position_with_budget(&candidate, patch.position(), &mut budget)?;
    if candidate_settings != *patch.after()
        || candidate.source_bytes() != patch.target_artifact().as_ref()
        || patch.touched_components() != 1
        || patch.source_layout_state().is_some()
        || patch.target_layout_state().is_some()
        || patch.source_preview_count() != patch.target_preview_count()
    {
        return Err(Error::Verification { path: patch.path() });
    }
    verify_unselected_members(
        source,
        &candidate,
        source_target,
        patch.before(),
        patch.after(),
    )?;
    budget.settle_transaction_reservation();
    Ok(Commit::from_parts(
        candidate,
        patch.clone(),
        Diagnostics::published(1, 0, true),
    ))
}

fn settings_at_position(package: &Package, position: Position) -> Result<Settings, Error> {
    let mut budget = transaction::TransactionBudget::new(package)?;
    settings_at_position_with_budget(package, position, &mut budget)
        .map(|(_target, settings)| settings)
}

fn settings_at_position_with_budget(
    package: &Package,
    position: Position,
    budget: &mut transaction::TransactionBudget,
) -> Result<(transaction::Target, Settings), Error> {
    let target = transaction::resolve_target(package, position, budget)?;
    let payload = transaction::selected_payload(package, target)?;
    let settings = decode_settings(package, payload, Path::section(position), budget)?;
    Ok((target, settings))
}

fn decode_settings(
    package: &Package,
    payload: &[u8],
    path: Path,
    budget: &mut transaction::TransactionBudget,
) -> Result<Settings, Error> {
    let limits = transaction::wire_limits(package)?;
    let options = DecodeOptions::new(
        limits.max_input_bytes(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    )
    .with_max_fields(budget.remaining_fields())
    .with_max_work_bytes(budget.remaining_work())
    .with_max_name_bytes(limits.max_input_bytes());
    let (snapshot, report) = decode_section_settings_with_report(payload, options)
        .map_err(|error| map_decode_error(&error, path))?;
    settle_report(report, budget, path)?;
    settings_from_snapshot(snapshot, path)
}

fn map_decode_error(
    error: &litchi_iwa_protos::pages_section_codec::DecodeError,
    path: Path,
) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource { path };
    };
    let (kind, observed, maximum) = match limit {
        DecodeLimit::Bytes { observed, maximum } => (
            crate::section::settings::LimitKind::WireInputBytes,
            transaction::usize_to_u64(observed),
            transaction::usize_to_u64(maximum),
        ),
        DecodeLimit::Fields { observed, maximum } => (
            crate::section::settings::LimitKind::WireFields,
            transaction::usize_to_u64(observed),
            transaction::usize_to_u64(maximum),
        ),
        DecodeLimit::Work { observed, maximum } => (
            crate::section::settings::LimitKind::WireWork,
            transaction::usize_to_u64(observed),
            transaction::usize_to_u64(maximum),
        ),
        DecodeLimit::Nesting { observed, maximum } => (
            crate::section::settings::LimitKind::WireNesting,
            u64::from(observed),
            u64::from(maximum),
        ),
        DecodeLimit::NameBytes { observed, maximum } => (
            crate::section::settings::LimitKind::RetainedBytes,
            transaction::usize_to_u64(observed),
            transaction::usize_to_u64(maximum),
        ),
        _ => return Error::InvalidSource { path },
    };
    Error::LimitExceeded {
        path,
        kind,
        observed,
        maximum,
    }
}

fn settle_report(
    report: DecodeReport,
    budget: &mut transaction::TransactionBudget,
    path: Path,
) -> Result<(), Error> {
    budget.charge_fields(report.fields(), path)?;
    budget.charge_work(report.work_bytes(), path)?;
    budget.charge_transaction_work(report.name_bytes(), path)
}

fn settings_from_snapshot(
    snapshot: SectionSettingsSnapshot<'_>,
    path: Path,
) -> Result<Settings, Error> {
    let mut settings = Settings::new();
    settings.set_inherit_previous_header_footer(snapshot.inherit_previous_header_footer());
    settings.set_first_page_different(snapshot.section_template_first_page_different());
    settings.set_even_odd_pages_different(snapshot.section_template_even_odd_pages_different());
    settings
        .set_start(snapshot.section_start_kind().map(Start::from_raw))
        .map_err(|_error| Error::InvalidSource { path })?;
    settings
        .set_page_numbering(
            snapshot
                .section_page_number_kind()
                .map(PageNumbering::from_raw),
        )
        .map_err(|_error| Error::InvalidSource { path })?;
    settings.set_starting_page_number(
        snapshot
            .section_page_number_start()
            .map(PageNumber::new)
            .transpose()
            .map_err(|_error| Error::InvalidSource { path })?,
    );
    let name = snapshot.name().map(try_boxed_str).transpose()?;
    settings
        .set_name(name)
        .map_err(|_error| Error::InvalidSource { path })?;
    settings.set_first_page_hides_header_footer(
        snapshot.section_template_first_page_hides_header_footer(),
    );
    settings.validate().map_err(Error::InvalidSettings)?;
    Ok(settings)
}

fn try_boxed_str(value: &str) -> Result<Box<str>, Error> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_error| Error::Allocation {
            amount: value.len(),
        })?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

fn rewrite_payload(
    package: &Package,
    source: &[u8],
    before: &Settings,
    after: &Settings,
    budget: &mut transaction::TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let edits = [
        NestedFieldEdit::new(
            &[INHERIT_FIELD],
            before.inherit_previous_header_footer().is_some(),
            NestedFieldReplacement::Varint(after.inherit_previous_header_footer().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &[FIRST_PAGE_DIFFERENT_FIELD],
            before.first_page_different().is_some(),
            NestedFieldReplacement::Varint(after.first_page_different().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &[EVEN_ODD_DIFFERENT_FIELD],
            before.even_odd_pages_different().is_some(),
            NestedFieldReplacement::Varint(after.even_odd_pages_different().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &[START_FIELD],
            before.start().is_some(),
            NestedFieldReplacement::Varint(after.start().map(|value| u64::from(value.as_raw()))),
        ),
        NestedFieldEdit::new(
            &[NUMBERING_FIELD],
            before.page_numbering().is_some(),
            NestedFieldReplacement::Varint(
                after
                    .page_numbering()
                    .map(|value| u64::from(value.as_raw())),
            ),
        ),
        NestedFieldEdit::new(
            &[STARTING_PAGE_FIELD],
            before.starting_page_number().is_some(),
            NestedFieldReplacement::Varint(
                after
                    .starting_page_number()
                    .map(|value| u64::from(value.get())),
            ),
        ),
        NestedFieldEdit::new(
            &[NAME_FIELD],
            before.name().is_some(),
            NestedFieldReplacement::LengthDelimited(after.name().map(str::as_bytes)),
        ),
        NestedFieldEdit::new(
            &[HIDE_FIRST_FIELD],
            before.first_page_hides_header_footer().is_some(),
            NestedFieldReplacement::Varint(after.first_page_hides_header_footer().map(u64::from)),
        ),
    ];
    let path = Path::Package;
    let rewritten =
        patch_nested_fields_batched_with_limits(source, &edits, transaction::wire_limits(package)?)
            .map_err(transaction::map_wire_error)?;
    budget.charge_work(
        source
            .len()
            .checked_add(rewritten.len())
            .ok_or(Error::LimitExceeded {
                path,
                kind: crate::section::settings::LimitKind::WireWork,
                observed: u64::MAX,
                maximum: u64::MAX - 1,
            })?,
        path,
    )?;
    if decode_settings(package, &rewritten, path, budget)? != *after {
        return Err(Error::Verification { path });
    }
    Ok(rewritten)
}

fn validate_changed_dependencies(
    package: &Package,
    target: transaction::Target,
    before: &Settings,
    after: &Settings,
    budget: &mut transaction::TransactionBudget,
) -> Result<(), Error> {
    let inherit_changed =
        before.inherit_previous_header_footer() != after.inherit_previous_header_footer();
    let first_changed = before.first_page_different() != after.first_page_different();
    let even_odd_changed = before.even_odd_pages_different() != after.even_odd_pages_different();
    let hide_changed =
        before.first_page_hides_header_footer() != after.first_page_hides_header_footer();
    if inherit_changed && target.position.get() == 0 {
        return Err(Error::UnsupportedDependency {
            path: Path::section(target.position),
            kind: crate::section::settings::DependencyKind::PreviousSectionTemplates,
        });
    }
    if !(inherit_changed || first_changed || even_odd_changed || hide_changed) {
        return Ok(());
    }
    let object_index = ObjectIndex::new(package, Path::section(target.position), budget)?;

    let fields: &[u32] = if inherit_changed || first_changed || hide_changed {
        &[
            FIRST_TEMPLATE_FIELD,
            EVEN_TEMPLATE_FIELD,
            ODD_TEMPLATE_FIELD,
        ]
    } else if even_odd_changed {
        &[EVEN_TEMPLATE_FIELD, ODD_TEMPLATE_FIELD]
    } else {
        &[]
    };
    let payload = transaction::selected_payload(package, target)?;
    for &field in fields {
        let kind = match field {
            FIRST_TEMPLATE_FIELD => crate::section::settings::DependencyKind::FirstTemplate,
            EVEN_TEMPLATE_FIELD => crate::section::settings::DependencyKind::EvenTemplate,
            _ => crate::section::settings::DependencyKind::OddTemplate,
        };
        let identifier =
            required_reference(package, payload, field, target.position, kind, budget)?;
        validate_selected_reference_metadata(package, target, field, identifier, budget)?;
        validate_template(
            package,
            &object_index,
            identifier,
            target.position,
            kind,
            budget,
        )?;
    }

    if inherit_changed && target.position.get() > 0 {
        let previous_position = Position::new(target.position.get() - 1);
        let previous = transaction::resolve_target(package, previous_position, budget)?;
        let previous_payload = transaction::selected_payload(package, previous)?;
        for (field, kind) in [
            (
                FIRST_TEMPLATE_FIELD,
                crate::section::settings::DependencyKind::FirstTemplate,
            ),
            (
                EVEN_TEMPLATE_FIELD,
                crate::section::settings::DependencyKind::EvenTemplate,
            ),
            (
                ODD_TEMPLATE_FIELD,
                crate::section::settings::DependencyKind::OddTemplate,
            ),
        ] {
            let identifier = required_reference(
                package,
                previous_payload,
                field,
                previous_position,
                kind,
                budget,
            )?;
            validate_selected_reference_metadata(package, previous, field, identifier, budget)?;
            validate_template(
                package,
                &object_index,
                identifier,
                previous_position,
                kind,
                budget,
            )?;
        }
    }
    Ok(())
}

fn required_reference(
    package: &Package,
    payload: &[u8],
    field_number: u32,
    position: Position,
    kind: crate::section::settings::DependencyKind,
    budget: &mut transaction::TransactionBudget,
) -> Result<u64, Error> {
    let path = Path::section(position);
    let view = WireView::parse_with_limits(payload, transaction::wire_limits(package)?)
        .map_err(transaction::map_wire_error)?;
    budget.charge_fields(view.len(), path)?;
    budget.charge_work(payload.len(), path)?;
    let mut selected = view.fields().filter(|field| field.number() == field_number);
    let Some(field) = selected.next() else {
        return Err(Error::UnsupportedDependency { path, kind });
    };
    if selected.next().is_some() || field.wire_type() != 2 {
        return Err(Error::InvalidSource { path });
    }
    budget.charge_references(1, path)?;
    let payload = field
        .canonical_payload()
        .map_err(transaction::map_wire_error)?;
    transaction::strict_reference(payload, transaction::wire_limits(package)?, path)
}

fn validate_template(
    package: &Package,
    object_index: &ObjectIndex<'_>,
    identifier: u64,
    position: Position,
    kind: crate::section::settings::DependencyKind,
    budget: &mut transaction::TransactionBudget,
) -> Result<(), Error> {
    let path = Path::section(position);
    let Some(object) = object_index.get(identifier, path, budget)? else {
        return Err(Error::UnsupportedDependency { path, kind });
    };
    let (message_index, message) =
        transaction::unique_message(object, transaction::TEMPLATE_MESSAGE_TYPE, path)?;
    transaction::validate_selected_metadata(object, message_index, path)?;
    budget.charge_work(message.data.len(), path)?;
    let view = WireView::parse_with_limits(&message.data, transaction::wire_limits(package)?)
        .map_err(transaction::map_wire_error)?;
    budget.charge_fields(view.len(), path)?;
    // Template drawables and UUID paths remain preservation-only. The native
    // Pages oracle retains both while changing only the selected section
    // scalar, so their presence is not a reason to reject the transaction.
    let mut headers = Vec::new();
    let mut footers = Vec::new();
    headers
        .try_reserve_exact(view.len())
        .map_err(|_error| Error::Allocation { amount: view.len() })?;
    footers
        .try_reserve_exact(view.len())
        .map_err(|_error| Error::Allocation { amount: view.len() })?;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(transaction::map_wire_error)?;
        if !matches!(field.number(), 1 | 2) {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource { path });
        }
        let storage = transaction::strict_reference(
            field
                .canonical_payload()
                .map_err(transaction::map_wire_error)?,
            transaction::wire_limits(package)?,
            path,
        )?;
        if field.number() == 1 {
            headers.push(storage);
        } else {
            footers.push(storage);
        }
        budget.charge_references(1, path)?;
        validate_storage(package, object_index, storage, path, kind, budget)?;
    }
    validate_repeated_reference_metadata(object, message_index, 1, &headers, path, budget)?;
    validate_repeated_reference_metadata(object, message_index, 2, &footers, path, budget)?;
    Ok(())
}

fn validate_repeated_reference_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
    field_number: u32,
    identifiers: &[u64],
    path: Path,
    budget: &mut transaction::TransactionBudget,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?;
    let metadata_references = info
        .field_infos
        .iter()
        .try_fold(0usize, |total, field| {
            total.checked_add(field.object_references.len())
        })
        .ok_or(Error::InvalidSource { path })?;
    let traversal_work = info
        .object_references
        .len()
        .saturating_add(info.field_infos.len())
        .saturating_add(metadata_references)
        .saturating_add(identifiers.len());
    let expected_sort_work = identifiers
        .len()
        .saturating_mul(ceil_log2(identifiers.len()));
    let aggregate_sort_work = info
        .object_references
        .len()
        .saturating_mul(ceil_log2(info.object_references.len()));
    budget.reserve_transaction_work(
        traversal_work
            .saturating_add(expected_sort_work)
            .saturating_add(aggregate_sort_work),
        path,
    )?;
    budget.charge_transaction_work(traversal_work, path)?;
    let mut expected = Vec::new();
    expected
        .try_reserve_exact(identifiers.len())
        .map_err(|_error| Error::Allocation {
            amount: identifiers.len(),
        })?;
    expected.extend_from_slice(identifiers);
    budget.charge_transaction_work(expected_sort_work, path)?;
    expected.sort_unstable();
    expected.dedup();
    let mut aggregate = Vec::new();
    aggregate
        .try_reserve_exact(info.object_references.len())
        .map_err(|_error| Error::Allocation {
            amount: info.object_references.len(),
        })?;
    aggregate.extend_from_slice(&info.object_references);
    budget.charge_transaction_work(aggregate_sort_work, path)?;
    aggregate.sort_unstable();
    for identifier in &expected {
        let start = aggregate.partition_point(|candidate| candidate < identifier);
        let end = aggregate.partition_point(|candidate| candidate <= identifier);
        if end.saturating_sub(start) != 1 {
            return Err(Error::InvalidSource { path });
        }
    }
    let mut declarations = info
        .field_infos
        .iter()
        .filter(|field| field.path.as_slice() == [field_number]);
    if let Some(declaration) = declarations.next()
        && (declarations.next().is_some()
            || declaration.object_references.as_slice() != identifiers
            || !declaration.data_references.is_empty())
    {
        return Err(Error::InvalidSource { path });
    }
    for field in &info.field_infos {
        if field.path.as_slice() == [field_number] {
            continue;
        }
        let aliases_known_role = matches!(field.path.as_slice(), [1] | [2]);
        if field
            .object_references
            .iter()
            .any(|identifier| expected.binary_search(identifier).is_ok())
            && (!aliases_known_role || !field.data_references.is_empty())
        {
            return Err(Error::InvalidSource { path });
        }
    }
    budget.settle_transaction_reservation();
    Ok(())
}

fn validate_storage(
    package: &Package,
    object_index: &ObjectIndex<'_>,
    identifier: u64,
    path: Path,
    kind: crate::section::settings::DependencyKind,
    budget: &mut transaction::TransactionBudget,
) -> Result<(), Error> {
    let object = object_index
        .get(identifier, path, budget)?
        .ok_or(Error::InvalidSource { path })?;
    if object.messages.len() != 1 {
        return Err(Error::InvalidSource { path });
    }
    let message_index = 0;
    let message = object
        .messages
        .first()
        .ok_or(Error::InvalidSource { path })?;
    if !transaction::STORAGE_MESSAGE_TYPES.contains(&message.type_) {
        return Err(Error::InvalidSource { path });
    }
    transaction::validate_selected_metadata(object, message_index, path)?;
    let view = WireView::parse_with_limits(&message.data, transaction::wire_limits(package)?)
        .map_err(transaction::map_wire_error)?;
    budget.charge_fields(view.len(), path)?;
    budget.charge_work(message.data.len(), path)?;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(transaction::map_wire_error)?;
        if matches!(
            field.number(),
            9 | 11 | 15 | 16 | 17 | 21 | 22 | 23 | 25 | 26
        ) {
            return Err(Error::UnsupportedDependency { path, kind });
        }
    }
    Ok(())
}

struct ObjectIndex<'a> {
    objects: Vec<(u64, &'a litchi_iwa_core::ArchiveObject)>,
}

impl<'a> ObjectIndex<'a> {
    fn new(
        package: &'a Package,
        path: Path,
        budget: &mut transaction::TransactionBudget,
    ) -> Result<Self, Error> {
        let components = package.state.source.components();
        budget.charge_transaction_work(components.len(), path)?;
        let count = components
            .iter()
            .try_fold(0usize, |total, component| {
                total.checked_add(component.archive().objects.len())
            })
            .ok_or(Error::InvalidSource { path })?;
        budget.charge_transaction_work(count, path)?;
        let mut objects = Vec::new();
        objects
            .try_reserve_exact(count)
            .map_err(|_error| Error::Allocation { amount: count })?;
        for component in components.iter() {
            for object in &component.archive().objects {
                let identifier = object
                    .archive_info
                    .identifier
                    .ok_or(Error::InvalidSource { path })?;
                objects.push((identifier, object));
            }
        }
        let sort_work = count.saturating_mul(ceil_log2(count));
        budget.charge_transaction_work(sort_work, path)?;
        objects.sort_unstable_by_key(|(identifier, _object)| *identifier);
        if objects.windows(2).any(|pair| pair[0].0 == pair[1].0) {
            return Err(Error::InvalidSource { path });
        }
        Ok(Self { objects })
    }

    fn get(
        &self,
        identifier: u64,
        path: Path,
        budget: &mut transaction::TransactionBudget,
    ) -> Result<Option<&'a litchi_iwa_core::ArchiveObject>, Error> {
        budget.charge_transaction_work(ceil_log2(self.objects.len()), path)?;
        Ok(self
            .objects
            .binary_search_by_key(&identifier, |(candidate, _object)| *candidate)
            .ok()
            .map(|index| self.objects[index].1))
    }
}

fn ceil_log2(value: usize) -> usize {
    if value <= 1 {
        0
    } else {
        usize::try_from(usize::BITS - (value - 1).leading_zeros()).unwrap_or(usize::MAX)
    }
}

fn validate_selected_reference_metadata(
    package: &Package,
    target: transaction::Target,
    field_number: u32,
    identifier: u64,
    budget: &mut transaction::TransactionBudget,
) -> Result<(), Error> {
    let path = Path::section(target.position);
    let component = package
        .state
        .source
        .components()
        .iter()
        .nth(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let object = component
        .archive()
        .objects
        .get(target.object_index)
        .ok_or(Error::InvalidSource { path })?;
    validate_reference_metadata(
        object,
        target.message_index,
        field_number,
        identifier,
        path,
        budget,
    )
}

fn validate_reference_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
    field_number: u32,
    identifier: u64,
    path: Path,
    budget: &mut transaction::TransactionBudget,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_transaction_work(
        info.object_references
            .len()
            .saturating_add(info.field_infos.len()),
        path,
    )?;
    if info
        .object_references
        .iter()
        .filter(|reference| **reference == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource { path });
    }
    let mut declarations = info
        .field_infos
        .iter()
        .filter(|field| field.path.as_slice() == [field_number]);
    if let Some(declaration) = declarations.next() {
        if declarations.next().is_some()
            || declaration.object_references.as_slice() != [identifier]
            || !declaration.data_references.is_empty()
        {
            return Err(Error::InvalidSource { path });
        }
    }
    for field in &info.field_infos {
        if field.path.as_slice() == [field_number] || !field.object_references.contains(&identifier)
        {
            continue;
        }
        if !matches!(field.path.as_slice(), [23] | [24] | [25])
            || field.object_references.as_slice() != [identifier]
            || !field.data_references.is_empty()
        {
            return Err(Error::InvalidSource { path });
        }
    }
    Ok(())
}

fn verify_unselected_members(
    source: &Package,
    candidate: &Package,
    target: transaction::Target,
    settings_before: &Settings,
    settings_after: &Settings,
) -> Result<(), Error> {
    let selected_name = source
        .state
        .source
        .components()
        .iter()
        .nth(target.component_index)
        .map(|component| component.name())
        .ok_or(Error::Verification {
            path: Path::Package,
        })?;
    let mut source_entries = source.state.source.package().iter();
    let mut candidate_entries = candidate.state.source.package().iter();
    loop {
        let (before, after) = match (source_entries.next(), candidate_entries.next()) {
            (Some(before), Some(after)) => (before, after),
            (None, None) => break,
            _ => {
                return Err(Error::Verification {
                    path: Path::Package,
                });
            },
        };
        if before.name() != after.name() {
            return Err(Error::Verification {
                path: Path::Package,
            });
        }
        let preserved = if before.name() == selected_name {
            selected_package_member_preserved(before, after)
        } else {
            package_member_preserved(before, after)
        };
        if !preserved {
            return Err(Error::Verification {
                path: Path::Package,
            });
        }
    }
    verify_selected_component(
        source,
        candidate,
        target,
        selected_name,
        settings_before,
        settings_after,
    )
}

fn verify_selected_component(
    source: &Package,
    candidate: &Package,
    target: transaction::Target,
    selected_name: &str,
    settings_before: &Settings,
    settings_after: &Settings,
) -> Result<(), Error> {
    let path = Path::section(target.position);
    let before_stream = selected_component_stream(source, selected_name)?;
    let after_stream = selected_component_stream(candidate, selected_name)?;
    let before_component = source
        .state
        .source
        .components()
        .get(selected_name)
        .ok_or(Error::Verification { path })?;
    let after_component = candidate
        .state
        .source
        .components()
        .get(selected_name)
        .ok_or(Error::Verification { path })?;
    if before_component.archive().objects.len() != after_component.archive().objects.len() {
        return Err(Error::Verification { path });
    }
    for (object_index, (before, after)) in before_component
        .archive()
        .objects
        .iter()
        .zip(&after_component.archive().objects)
        .enumerate()
    {
        if object_index != target.object_index {
            if before.archive_info != after.archive_info
                || before.messages != after.messages
                || raw_object(before_stream.as_bytes(), before)
                    != raw_object(after_stream.as_bytes(), after)
            {
                return Err(Error::Verification { path });
            }
            continue;
        }
        let before_header = raw_object_header(before_stream.as_bytes(), before)
            .ok_or(Error::Verification { path })?;
        let after_header = raw_object_header(after_stream.as_bytes(), after)
            .ok_or(Error::Verification { path })?;
        if normalized_message_length(before_header, target.message_index, source)?
            != normalized_message_length(after_header, target.message_index, candidate)?
        {
            return Err(Error::Verification { path });
        }
        if before.archive_info.identifier != after.archive_info.identifier
            || before.archive_info.should_merge != after.archive_info.should_merge
            || before.messages.len() != after.messages.len()
            || before.archive_info.message_infos.len() != after.archive_info.message_infos.len()
        {
            return Err(Error::Verification { path });
        }
        for (message_index, (old, new)) in before.messages.iter().zip(&after.messages).enumerate() {
            let old_info = before
                .archive_info
                .message_infos
                .get(message_index)
                .ok_or(Error::Verification { path })?;
            let new_info = after
                .archive_info
                .message_infos
                .get(message_index)
                .ok_or(Error::Verification { path })?;
            if message_index == target.message_index {
                if old.type_ != new.type_
                    || !message_info_preserved_except_length(old_info, new_info)
                    || !unchanged_section_fields_preserved(
                        &old.data,
                        &new.data,
                        source,
                        settings_before,
                        settings_after,
                    )?
                {
                    return Err(Error::Verification { path });
                }
            } else if old != new || old_info != new_info {
                return Err(Error::Verification { path });
            }
        }
    }
    Ok(())
}

fn selected_component_stream(package: &Package, name: &str) -> Result<SnappyStream, Error> {
    let entry = package
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(Error::Verification {
            path: Path::Package,
        })?;
    if entry.is_opaque() {
        return Err(Error::Verification {
            path: Path::Package,
        });
    }
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        package
            .state
            .source
            .limits()
            .snappy_limits()
            .map_err(transaction::map_archive_error)?,
    )
    .map_err(transaction::map_core_error)?;
    Ok(stream)
}

fn raw_object<'a>(stream: &'a [u8], object: &litchi_iwa_core::ArchiveObject) -> Option<&'a [u8]> {
    let start = usize::try_from(object.header_offset).ok()?;
    let end = usize::try_from(object.data_offset)
        .ok()?
        .checked_add(usize::try_from(object.data_length).ok()?)?;
    stream.get(start..end)
}

fn raw_object_header<'a>(
    stream: &'a [u8],
    object: &litchi_iwa_core::ArchiveObject,
) -> Option<&'a [u8]> {
    let offset = usize::try_from(object.header_offset).ok()?;
    let (_, prefix) = decode_varint_from_bytes(stream.get(offset..)?).ok()?;
    let start = offset.checked_add(prefix)?;
    let end = usize::try_from(object.data_offset).ok()?;
    stream.get(start..end)
}

fn normalized_message_length(
    header: &[u8],
    target_index: usize,
    package: &Package,
) -> Result<Vec<u8>, Error> {
    let view = WireView::parse_with_limits(header, transaction::wire_limits(package)?)
        .map_err(transaction::map_wire_error)?;
    let reserve = header.len().saturating_add(48);
    let mut output = Vec::new();
    output
        .try_reserve_exact(reserve)
        .map_err(|_error| Error::Allocation { amount: reserve })?;
    let mut message_index = 0usize;
    for field in view.fields() {
        if field.number() != 2 || field.wire_type() != 2 {
            output.extend_from_slice(field.raw());
            continue;
        }
        if message_index != target_index {
            output.extend_from_slice(field.raw());
            message_index = message_index.saturating_add(1);
            continue;
        }
        output.extend_from_slice(field.key());
        output.extend_from_slice(b"<message-info>");
        let message =
            WireView::parse_with_limits(field.payload(), transaction::wire_limits(package)?)
                .map_err(transaction::map_wire_error)?;
        let effective_length = message
            .fields()
            .enumerate()
            .filter_map(|(index, nested)| {
                (nested.number() == 3 && nested.wire_type() == 0).then_some(index)
            })
            .last()
            .ok_or(Error::Verification {
                path: Path::Package,
            })?;
        for (index, nested) in message.fields().enumerate() {
            if index == effective_length {
                output.extend_from_slice(nested.key());
                output.extend_from_slice(b"<payload-length>");
            } else {
                output.extend_from_slice(nested.raw());
            }
        }
        output.extend_from_slice(b"</message-info>");
        message_index = message_index.saturating_add(1);
    }
    Ok(output)
}

fn unchanged_section_fields_preserved(
    source: &[u8],
    candidate: &[u8],
    package: &Package,
    before: &Settings,
    after: &Settings,
) -> Result<bool, Error> {
    let limits = transaction::wire_limits(package)?;
    let left = WireView::parse_with_limits(source, limits).map_err(transaction::map_wire_error)?;
    let right =
        WireView::parse_with_limits(candidate, limits).map_err(transaction::map_wire_error)?;
    let mut left_fields = left
        .fields()
        .filter(|field| !owned_field_changed(field.number(), before, after));
    let mut right_fields = right
        .fields()
        .filter(|field| !owned_field_changed(field.number(), before, after));
    loop {
        match (left_fields.next(), right_fields.next()) {
            (Some(left), Some(right)) if left.raw() == right.raw() => {},
            (None, None) => return Ok(true),
            _ => return Ok(false),
        }
    }
}

fn owned_field_changed(number: u32, before: &Settings, after: &Settings) -> bool {
    match number {
        INHERIT_FIELD => {
            before.inherit_previous_header_footer() != after.inherit_previous_header_footer()
        },
        FIRST_PAGE_DIFFERENT_FIELD => before.first_page_different() != after.first_page_different(),
        EVEN_ODD_DIFFERENT_FIELD => {
            before.even_odd_pages_different() != after.even_odd_pages_different()
        },
        START_FIELD => before.start() != after.start(),
        NUMBERING_FIELD => before.page_numbering() != after.page_numbering(),
        STARTING_PAGE_FIELD => before.starting_page_number() != after.starting_page_number(),
        NAME_FIELD => before.name() != after.name(),
        HIDE_FIRST_FIELD => {
            before.first_page_hides_header_footer() != after.first_page_hides_header_footer()
        },
        _ => false,
    }
}

fn message_info_preserved_except_length(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && source.object_references == candidate.object_references
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

fn package_member_preserved(source: &Entry, candidate: &Entry) -> bool {
    source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_package_member_preserved(source: &Entry, candidate: &Entry) -> bool {
    source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.metadata().local() == candidate.metadata().local()
        && source.metadata().central() == candidate.metadata().central()
        && selected_local_record_preserved(source, candidate)
        && selected_central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..OFFSET.start] == candidate[..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}

fn selected_local_record_preserved(source: &Entry, candidate: &Entry) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 14..26;
    let left = source.raw_record().local_record();
    let right = candidate.raw_record().local_record();
    let (Some(left_header), Some(right_header)) = (
        zip_local_header_length(left),
        zip_local_header_length(right),
    ) else {
        return false;
    };
    if left_header != right_header
        || left[..CRC_AND_SIZES.start] != right[..CRC_AND_SIZES.start]
        || left[CRC_AND_SIZES.end..left_header] != right[CRC_AND_SIZES.end..right_header]
    {
        return false;
    }
    let Some(left_end) = left_header
        .checked_add(source.raw_record().compressed_data().len())
        .filter(|end| *end <= left.len())
    else {
        return false;
    };
    let Some(right_end) = right_header
        .checked_add(candidate.raw_record().compressed_data().len())
        .filter(|end| *end <= right.len())
    else {
        return false;
    };
    selected_local_suffix_preserved(
        source.metadata().local().flags(),
        &left[left_end..],
        &right[right_end..],
    )
}

fn zip_local_header_length(record: &[u8]) -> Option<usize> {
    if record.get(..4)? != b"PK\x03\x04" {
        return None;
    }
    let name = usize::from(u16::from_le_bytes(record.get(26..28)?.try_into().ok()?));
    let extra = usize::from(u16::from_le_bytes(record.get(28..30)?.try_into().ok()?));
    30usize
        .checked_add(name)?
        .checked_add(extra)
        .filter(|length| *length <= record.len())
}

fn selected_local_suffix_preserved(flags: u16, source: &[u8], candidate: &[u8]) -> bool {
    if flags & 0x0008 == 0 {
        return source == candidate;
    }
    let source_prefix = usize::from(source.starts_with(b"PK\x07\x08")) * 4;
    let candidate_prefix = usize::from(candidate.starts_with(b"PK\x07\x08")) * 4;
    source_prefix == candidate_prefix
        && source.len() == candidate.len()
        && source.len() >= source_prefix + 12
        && source[..source_prefix] == candidate[..candidate_prefix]
        && source[source_prefix + 12..] == candidate[candidate_prefix + 12..]
}

fn selected_central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 16..28;
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..CRC_AND_SIZES.start] == candidate[..CRC_AND_SIZES.start]
        && source[CRC_AND_SIZES.end..OFFSET.start] == candidate[CRC_AND_SIZES.end..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}

#[cfg(test)]
pub(super) fn production_test_usage(
    source: &Package,
    after: &Settings,
    transaction_limit: Option<usize>,
) -> Result<transaction::Usage, Error> {
    let mut budget = match transaction_limit {
        Some(maximum) => transaction::TransactionBudget::with_transaction_limit(source, maximum)?,
        None => transaction::TransactionBudget::new(source)?,
    };
    production_test_run(source, after, &mut budget)?;
    Ok(budget.usage())
}

#[cfg(test)]
pub(super) fn production_test_attempt(
    source: &Package,
    after: &Settings,
    transaction_limit: usize,
) -> (Result<(), Error>, transaction::Usage) {
    let mut budget =
        match transaction::TransactionBudget::with_transaction_limit(source, transaction_limit) {
            Ok(budget) => budget,
            Err(error) => return (Err(error), transaction::Usage::default()),
        };
    let result = production_test_run(source, after, &mut budget);
    (result, budget.usage())
}

#[cfg(test)]
fn production_test_run(
    source: &Package,
    after: &Settings,
    budget: &mut transaction::TransactionBudget,
) -> Result<(), Error> {
    let position = Position::new(0);
    let (target, before) = settings_at_position_with_budget(source, position, budget)?;
    validate_changed_dependencies(source, target, &before, after, budget)?;
    let payload = transaction::selected_payload(source, target)?;
    let rewritten = rewrite_payload(source, payload, &before, after, budget)?;
    let (candidate, stats) =
        transaction::rewrite_package(source, target, rewritten, false, budget)?;
    if stats.touched_components != 1 || stats.deleted_previews != 0 {
        return Err(Error::Verification {
            path: Path::Package,
        });
    }
    let (_candidate_target, observed) =
        settings_at_position_with_budget(&candidate, position, budget)?;
    if observed != *after {
        return Err(Error::Verification {
            path: Path::Package,
        });
    }
    verify_unselected_members(source, &candidate, target, &before, after)?;
    budget.settle_transaction_reservation();
    Ok(())
}

#[cfg(test)]
mod tests {
    use litchi_core::Position;
    use litchi_iwa_core::{ArchiveObject, FieldInfo, RawMessage};

    use super::{
        Error, Package, Path, transaction, validate_reference_metadata,
        validate_repeated_reference_metadata,
    };

    type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

    fn package() -> TestResult<Package> {
        Ok(Package::from_bytes(include_bytes!(
            "../../../../test-data/iwork/pages/basic.pages"
        ))?)
    }

    fn metadata_object(
        aggregate: Vec<u64>,
        declarations: impl IntoIterator<Item = (u32, Vec<u64>)>,
    ) -> TestResult<ArchiveObject> {
        let mut object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 10_143,
                data: Vec::new(),
            }],
        )?;
        object.archive_info.message_infos[0].object_references = aggregate;
        for (field_number, references) in declarations {
            let mut field = FieldInfo::new(vec![field_number]);
            field.object_references = references;
            object.archive_info.message_infos[0].field_infos.push(field);
        }
        Ok(object)
    }

    #[test]
    fn aliased_template_and_header_footer_roles_accept_complete_metadata() -> TestResult {
        let package = package()?;
        let path = Path::section(Position::new(1));
        let mut budget = transaction::TransactionBudget::new(&package)?;
        let section = metadata_object(vec![77], [(23, vec![77]), (24, vec![77]), (25, vec![77])])?;
        for field in [23, 24, 25] {
            validate_reference_metadata(&section, 0, field, 77, path, &mut budget)?;
        }

        let template = metadata_object(vec![88], [(1, vec![88]), (2, vec![88])])?;
        validate_repeated_reference_metadata(&template, 0, 1, &[88], path, &mut budget)?;
        validate_repeated_reference_metadata(&template, 0, 2, &[88], path, &mut budget)?;
        Ok(())
    }

    fn repeated_metadata_attempt(
        count: usize,
        maximum: Option<usize>,
    ) -> TestResult<(Result<(), Error>, transaction::Usage)> {
        let package = package()?;
        let identifiers = (0..count)
            .map(|index| 10_000_u64.saturating_add(u64::try_from(index).unwrap_or(u64::MAX)))
            .collect::<Vec<_>>();
        let object = metadata_object(identifiers.clone(), [(1, identifiers.clone())])?;
        let mut budget = match maximum {
            Some(maximum) => {
                transaction::TransactionBudget::with_transaction_limit(&package, maximum)?
            },
            None => transaction::TransactionBudget::new(&package)?,
        };
        let result = validate_repeated_reference_metadata(
            &object,
            0,
            1,
            &identifiers,
            Path::section(Position::new(1)),
            &mut budget,
        );
        Ok((result, budget.usage()))
    }

    #[test]
    fn repeated_reference_metadata_scales_and_max_minus_one_precharges() -> TestResult {
        let (small_result, small) = repeated_metadata_attempt(4_096, None)?;
        small_result?;
        let (large_result, large) = repeated_metadata_attempt(8_192, None)?;
        large_result?;
        assert!(small.transaction_work != 0);
        assert!(
            large.transaction_work.saturating_mul(10)
                <= small.transaction_work.saturating_mul(23) + 320
        );

        let (attempt, usage) = repeated_metadata_attempt(8_192, Some(large.transaction_work - 1))?;
        assert!(matches!(
            attempt,
            Err(Error::LimitExceeded {
                kind: crate::section::settings::LimitKind::TransactionWork,
                ..
            })
        ));
        assert_eq!(usage.transaction_work, 0);
        Ok(())
    }
}
