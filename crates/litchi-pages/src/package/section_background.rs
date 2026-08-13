//! Exact-source transactions for the fill stored directly on one Pages section.

use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_common::{WireLimits, color::RgbColorSpace, wire::WireView};
use litchi_iwa_protos::pages_section_background_codec::{
    self as codec, BackgroundSnapshot, BackgroundWrite, DecodeOptions, RgbSpace,
};

use super::{Package, section_transaction as transaction};
use crate::{
    SectionSelector,
    section::{
        Background,
        background::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
    },
};

impl Package {
    /// Read the semantic fill stored directly on one selected section.
    pub fn section_background<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Background, Error> {
        let position = transaction::resolve_position(self, selector).map_err(map_transaction)?;
        background_at_position(self, position)
    }

    /// Begin one immutable section-background transaction.
    pub fn edit_section_background<'package, 'selector>(
        &'package self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Edit<'package>, Error> {
        let position = transaction::resolve_position(self, selector).map_err(map_transaction)?;
        let before = background_at_position(self, position)?;
        Ok(Edit::from_parts(self, position, before))
    }

    /// Apply a reversible background patch to this exact package artifact.
    pub fn apply_section_background(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}

pub(crate) fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let (source, position, before, after) = edit.into_parts();
    let source_artifact = source.state.source.shared_source();
    let source_fingerprint = transaction::fingerprint(source_artifact.as_ref());
    if before == after {
        let patch = Patch::from_parts(
            Arc::clone(&source_artifact),
            source_artifact,
            source_fingerprint,
            source_fingerprint,
            position,
            before,
            after,
            0,
        );
        return Ok(Commit::from_parts(
            source.snapshot(),
            patch,
            Diagnostics::unchanged(),
        ));
    }
    if before == Background::Unsupported {
        return Err(Error::UnsupportedSource {
            path: Path::section(position),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(Error::UnsupportedSource {
            path: Path::section(position),
        });
    }

    let mut budget = transaction::TransactionBudget::new(source).map_err(map_transaction)?;
    let target =
        transaction::resolve_target(source, position, &mut budget).map_err(map_transaction)?;
    validate_reference_ownership(source, target, &mut budget)?;
    let payload = transaction::selected_payload(source, target).map_err(map_transaction)?;
    let decoded = decode_background(source, payload, Path::section(position), &mut budget)?;
    if decoded != before {
        return Err(Error::InvalidSource {
            path: Path::section(position),
        });
    }
    let rewritten = rewrite_background(source, payload, &after, &mut budget)?;
    let (candidate, stats) =
        transaction::rewrite_package(source, target, rewritten, false, &mut budget)
            .map_err(map_transaction)?;
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
    let (_candidate_target, candidate_background) =
        background_at_position_with_budget(&candidate, position, &mut budget)?;
    if candidate_background != after {
        return Err(Error::Verification {
            path: Path::section(position),
        });
    }
    verify_locality(source, &candidate, target)?;
    budget.settle_transaction_reservation();

    let target_artifact = candidate.state.source.shared_source();
    let target_fingerprint = transaction::fingerprint(target_artifact.as_ref());
    let patch = Patch::from_parts(
        source_artifact,
        target_artifact,
        source_fingerprint,
        target_fingerprint,
        position,
        before,
        after,
        stats.touched_components,
    );
    Ok(Commit::from_parts(
        candidate,
        patch,
        Diagnostics::published(stats.touched_components),
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
    let mut budget = transaction::TransactionBudget::new(source).map_err(map_transaction)?;
    let (source_target, source_background) =
        background_at_position_with_budget(source, patch.position(), &mut budget)?;
    if source_background != *patch.before() || source_background == Background::Unsupported {
        return Err(Error::PatchConflict);
    }
    validate_reference_ownership(source, source_target, &mut budget)?;
    budget
        .charge_transaction_work(
            patch
                .target_artifact()
                .len()
                .saturating_mul(4)
                .saturating_add(128),
            transaction_path(Path::Package),
        )
        .map_err(map_transaction)?;
    let catalog = SourceCatalog::from_shared_bytes_with_limits(
        Arc::clone(patch.target_artifact()),
        source.state.source.limits(),
    )
    .map_err(transaction::map_archive_error)
    .map_err(map_transaction)?;
    let candidate = Package::from_source_catalog(catalog)
        .map_err(transaction::map_package_error)
        .map_err(map_transaction)?;
    let (_candidate_target, candidate_background) =
        background_at_position_with_budget(&candidate, patch.position(), &mut budget)?;
    if candidate_background != *patch.after()
        || candidate.source_bytes() != patch.target_artifact().as_ref()
        || patch.touched_components() != 1
    {
        return Err(Error::Verification { path: patch.path() });
    }
    verify_locality(source, &candidate, source_target)?;
    budget.settle_transaction_reservation();
    Ok(Commit::from_parts(
        candidate,
        patch.clone(),
        Diagnostics::published(1),
    ))
}

fn background_at_position(package: &Package, position: Position) -> Result<Background, Error> {
    let mut budget = transaction::TransactionBudget::new(package).map_err(map_transaction)?;
    background_at_position_with_budget(package, position, &mut budget)
        .map(|(_target, background)| background)
}

fn background_at_position_with_budget(
    package: &Package,
    position: Position,
    budget: &mut transaction::TransactionBudget,
) -> Result<(transaction::Target, Background), Error> {
    let target = transaction::resolve_target(package, position, budget).map_err(map_transaction)?;
    let payload = transaction::selected_payload(package, target).map_err(map_transaction)?;
    let background = decode_background(package, payload, Path::section(position), budget)?;
    Ok((target, background))
}

fn decode_background(
    package: &Package,
    payload: &[u8],
    path: Path,
    budget: &mut transaction::TransactionBudget,
) -> Result<Background, Error> {
    let limits = transaction::wire_limits(package).map_err(map_transaction)?;
    let options = codec_options(limits, budget);
    let (snapshot, report) = codec::decode_section_background_with_report(payload, options)
        .map_err(|error| map_codec(&error, path))?;
    settle_decode(report.fields, report.work_bytes, budget, path)?;
    snapshot_to_background(snapshot, path)
}

fn rewrite_background(
    package: &Package,
    payload: &[u8],
    after: &Background,
    budget: &mut transaction::TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let limits = transaction::wire_limits(package).map_err(map_transaction)?;
    let write = match after {
        Background::None => BackgroundWrite::Clear,
        Background::Solid(color) => BackgroundWrite::Solid {
            red: color.red(),
            green: color.green(),
            blue: color.blue(),
            alpha: color.alpha(),
            rgb_space: match color.color_space() {
                RgbColorSpace::Srgb => RgbSpace::Srgb,
                RgbColorSpace::DisplayP3 => RgbSpace::DisplayP3,
            },
        },
        Background::Unsupported => {
            return Err(Error::UnsupportedSource {
                path: Path::Package,
            });
        },
    };
    let options = codec_options(limits, budget);
    let (rewritten, report) =
        codec::rewrite_section_background_with_report(payload, write, options)
            .map_err(|error| map_codec(&error, Path::Package))?;
    settle_decode(report.fields, report.work_bytes, budget, Path::Package)?;
    Ok(rewritten)
}

fn codec_options(limits: WireLimits, budget: &transaction::TransactionBudget) -> DecodeOptions {
    DecodeOptions::new(
        limits.max_input_bytes(),
        limits.max_output_bytes(),
        budget.remaining_fields(),
        budget.remaining_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    )
}

fn settle_decode(
    fields: usize,
    work: usize,
    budget: &mut transaction::TransactionBudget,
    path: Path,
) -> Result<(), Error> {
    budget
        .charge_fields(fields, transaction_path(path))
        .and_then(|()| budget.charge_work(work, transaction_path(path)))
        .map_err(map_transaction)
}

fn snapshot_to_background(snapshot: BackgroundSnapshot, path: Path) -> Result<Background, Error> {
    match snapshot {
        BackgroundSnapshot::None => Ok(Background::None),
        BackgroundSnapshot::Unsupported => Ok(Background::Unsupported),
        BackgroundSnapshot::Solid {
            red,
            green,
            blue,
            alpha,
            rgb_space,
        } => litchi_iwa_common::color::Rgba::new(
            red,
            green,
            blue,
            alpha,
            match rgb_space {
                RgbSpace::Srgb => RgbColorSpace::Srgb,
                RgbSpace::DisplayP3 => RgbColorSpace::DisplayP3,
            },
        )
        .map(Background::Solid)
        .map_err(|_error| Error::InvalidSource { path }),
    }
}

fn validate_reference_ownership(
    package: &Package,
    target: transaction::Target,
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
    let info = object
        .archive_info
        .message_infos
        .get(target.message_index)
        .ok_or(Error::InvalidSource { path })?;
    let references = info
        .field_infos
        .iter()
        .try_fold(
            info.object_references
                .len()
                .saturating_add(info.data_references.len()),
            |total, field| {
                total
                    .checked_add(field.object_references.len())
                    .and_then(|value| value.checked_add(field.data_references.len()))
            },
        )
        .ok_or(Error::InvalidSource { path })?;
    budget
        .charge_references(references, transaction_path(path))
        .map_err(map_transaction)?;
    if !info.data_references.is_empty()
        || info.field_infos.iter().any(|field| {
            field.path.path.first() == Some(&30)
                && (!field.object_references.is_empty() || !field.data_references.is_empty())
        })
    {
        return Err(Error::InvalidSource { path });
    }
    let payload = transaction::selected_payload(package, target).map_err(map_transaction)?;
    let view = WireView::parse_with_limits(
        payload,
        transaction::wire_limits(package).map_err(map_transaction)?,
    )
    .map_err(transaction::map_wire_error)
    .map_err(map_transaction)?;
    let mut preserved = [None; 4];
    for field in view.fields() {
        let slot = match field.number() {
            23 => 0,
            24 => 1,
            25 => 2,
            29 => 3,
            _ => continue,
        };
        if preserved[slot].is_some() || field.wire_type() != 2 {
            return Err(Error::InvalidSource { path });
        }
        preserved[slot] = Some(
            transaction::strict_reference(
                field
                    .canonical_payload()
                    .map_err(transaction::map_wire_error)
                    .map_err(map_transaction)?,
                transaction::wire_limits(package).map_err(map_transaction)?,
                transaction_path(path),
            )
            .map_err(map_transaction)?,
        );
    }
    let mut expected = [0_u64; 4];
    let mut expected_count = 0_usize;
    for identifier in preserved.into_iter().flatten() {
        if !expected[..expected_count].contains(&identifier) {
            expected[expected_count] = identifier;
            expected_count += 1;
        }
    }
    if info.object_references.len() != expected_count
        || expected[..expected_count].iter().any(|identifier| {
            info.object_references
                .iter()
                .filter(|candidate| *candidate == identifier)
                .count()
                != 1
        })
        || info
            .object_references
            .iter()
            .any(|identifier| !expected[..expected_count].contains(identifier))
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    target: transaction::Target,
) -> Result<(), Error> {
    super::section_settings::verify_unselected_members(
        source,
        candidate,
        target,
        &crate::section::Settings::new(),
        &crate::section::Settings::new(),
        Some(30),
    )
    .map_err(map_transaction)
}

fn map_codec(error: &codec::Error, path: Path) -> Error {
    let Some(limit) = error.limit() else {
        return Error::InvalidSource { path };
    };
    match limit {
        codec::LimitKind::InputBytes { observed, maximum } => {
            limit_error(path, LimitKind::WireInputBytes, observed, maximum)
        },
        codec::LimitKind::OutputBytes { observed, maximum } => {
            limit_error(path, LimitKind::WireOutputBytes, observed, maximum)
        },
        codec::LimitKind::Fields { observed, maximum } => {
            limit_error(path, LimitKind::WireFields, observed, maximum)
        },
        codec::LimitKind::WorkBytes { observed, maximum } => {
            limit_error(path, LimitKind::WireWork, observed, maximum)
        },
        codec::LimitKind::Nesting { observed, maximum } => Error::LimitExceeded {
            path,
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        codec::LimitKind::Allocation { requested } => Error::Allocation { amount: requested },
        _ => Error::InvalidSource { path },
    }
}

fn limit_error(path: Path, kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::LimitExceeded {
        path,
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    }
}

fn transaction_path(path: Path) -> crate::section::settings::Path {
    match path {
        Path::Package => crate::section::settings::Path::Package,
        Path::Section { position } => crate::section::settings::Path::section(position),
    }
}

fn map_transaction(error: crate::section::settings::Error) -> Error {
    use crate::section::settings::{Error as Source, LimitKind as SourceLimit};
    match error {
        Source::AmbiguousSelector { first, duplicate } => {
            Error::AmbiguousSelector { first, duplicate }
        },
        Source::NameNotFound => Error::NameNotFound,
        Source::PositionNotFound { position } => Error::PositionNotFound { position },
        Source::UnsupportedSource { path } | Source::UnsupportedDependency { path, .. } => {
            Error::UnsupportedSource {
                path: background_path(path),
            }
        },
        Source::InvalidSource { path } => Error::InvalidSource {
            path: background_path(path),
        },
        Source::InvalidSettings(_) => Error::InvalidSource {
            path: Path::Package,
        },
        Source::LimitExceeded {
            path,
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            path: background_path(path),
            kind: match kind {
                SourceLimit::InputBytes => LimitKind::InputBytes,
                SourceLimit::OutputBytes => LimitKind::OutputBytes,
                SourceLimit::Entries => LimitKind::Entries,
                SourceLimit::EntryBytes => LimitKind::EntryBytes,
                SourceLimit::TotalEntryBytes => LimitKind::TotalEntryBytes,
                SourceLimit::PackageBytes => LimitKind::PackageBytes,
                SourceLimit::PayloadBytes => LimitKind::PayloadBytes,
                SourceLimit::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
                SourceLimit::PayloadObjects => LimitKind::PayloadObjects,
                SourceLimit::PayloadMessages => LimitKind::PayloadMessages,
                SourceLimit::PayloadItems => LimitKind::PayloadItems,
                SourceLimit::References => LimitKind::References,
                SourceLimit::RetainedBytes => LimitKind::RetainedBytes,
                SourceLimit::WireInputBytes => LimitKind::WireInputBytes,
                SourceLimit::WireOutputBytes => LimitKind::WireOutputBytes,
                SourceLimit::WireFields => LimitKind::WireFields,
                SourceLimit::WireNesting => LimitKind::WireNesting,
                SourceLimit::WireWork => LimitKind::WireWork,
                SourceLimit::TransactionWork => LimitKind::TransactionWork,
            },
            observed,
            maximum,
        },
        Source::Allocation { amount } => Error::Allocation { amount },
        Source::Verification { path } => Error::Verification {
            path: background_path(path),
        },
        Source::PatchConflict => Error::PatchConflict,
    }
}

fn background_path(path: crate::section::settings::Path) -> Path {
    match path {
        crate::section::settings::Path::Package => Path::Package,
        crate::section::settings::Path::Section { position } => Path::section(position),
    }
}

#[cfg(test)]
mod perf_tests;

#[cfg(test)]
pub(super) fn production_test_attempt(
    package: &Package,
    requested: Background,
    maximum_transaction_work: Option<usize>,
) -> (Result<(), Error>, transaction::Usage) {
    let mut budget = maximum_transaction_work.map_or_else(
        || transaction::TransactionBudget::new(package),
        |maximum| transaction::TransactionBudget::with_transaction_limit(package, maximum),
    );
    let result = budget
        .as_mut()
        .map_err(|error| map_transaction(*error))
        .and_then(|budget| {
            let position = Position::new(0);
            let target =
                transaction::resolve_target(package, position, budget).map_err(map_transaction)?;
            validate_reference_ownership(package, target, budget)?;
            let payload =
                transaction::selected_payload(package, target).map_err(map_transaction)?;
            let rewritten = rewrite_background(package, payload, &requested, budget)?;
            let (candidate, stats) =
                transaction::rewrite_package(package, target, rewritten, false, budget)
                    .map_err(map_transaction)?;
            if stats.touched_components != 1 || stats.deleted_previews != 0 {
                return Err(Error::Verification {
                    path: Path::Package,
                });
            }
            let (_candidate_target, observed) =
                background_at_position_with_budget(&candidate, position, budget)?;
            if observed != requested {
                return Err(Error::Verification {
                    path: Path::Package,
                });
            }
            verify_locality(package, &candidate, target)?;
            budget.settle_transaction_reservation();
            Ok(())
        });
    let usage = budget.map_or_else(
        |_error| transaction::Usage::default(),
        transaction::TransactionBudget::usage,
    );
    (result, usage)
}
