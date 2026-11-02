//! Exact-source publication of a Pop-Up Menu change (lifted from the
//! Numbers Pop-Up Menu facade, record 0412a).
//!
//! The native graph rewrite, the metadata registry transition and the menu
//! model allocation are this crate's [`super::pop_up_menu`] engine; this
//! module stages their members, the package metadata member and (when the
//! leaf invalidates them) the root preview deletions as one exact-source
//! reassembly, reopens the candidate through the leaf, proves package
//! locality and reads the menu back. Numbers, Pages and Keynote publish
//! their Pop-Up Menu writes through it.

use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::table::cell::data_format::PopUpMenu;

use super::pop_up_menu::{
    map_popup_metadata_error, pack_metadata_payload, popup_metadata_options,
    read_popup_with_budget, rewrite_native_entry, rewrite_popup_metadata,
};
use super::{Error, ListKeys, Path, Refusal, Target, TransactionBudget, pop_up_menu_metadata};
use crate::TableSource;

/// The package metadata member.
const METADATA_ENTRY: &str = "Index/Metadata.iwa";

/// One published Pop-Up Menu candidate.
#[derive(Debug)]
pub struct Publication<P> {
    /// The reopened candidate package.
    pub package: P,
    /// Native components and the metadata component the edit rewrote.
    pub touched_components: usize,
    /// Root previews the publication deleted.
    pub deleted_previews: usize,
}

/// Publish a Pop-Up Menu change of the selected stored cell: `before` is
/// the menu the edit observed (a conflict otherwise), `after` the menu to
/// store (`None` clears it). `delete_previews` selects the leaf's root
/// preview policy, `list_keys` its format-list key policy and
/// `register_model_uuid` whether a new menu model's UUID is registered; `reopen` reopens the owned candidate through the leaf's
/// strict ingress and `retarget` resolves the same cell on it.
///
/// # Errors
///
/// The engine's typed refusals, limits, allocation and verification
/// failures; [`Error::PatchConflict`] when the cell no longer holds
/// `before`.
#[allow(
    clippy::too_many_lines,
    reason = "one staged publication: native graph, metadata, reassembly, reopen, locality, readback"
)]
pub fn publish<P: TableSource>(
    source: &P,
    target: Target,
    before: Option<&PopUpMenu>,
    after: Option<&PopUpMenu>,
    delete_previews: bool,
    list_keys: ListKeys,
    register_model_uuid: bool,
    reopen: impl FnOnce(Vec<u8>) -> Result<P, Error>,
    retarget: impl FnOnce(&P) -> Result<Target, Error>,
) -> Result<Publication<P>, Error> {
    let path = target.path();
    if target.locked {
        return Err(Error::Refused(Refusal::TableLocked { path }));
    }
    let mut budget = TransactionBudget::new(source).with_list_keys(list_keys);
    let catalog = source
        .physical()
        .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    if !catalog.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    budget.charge_package_source(catalog, path)?;
    let observed = read_popup_with_budget(source, target, &mut budget, false)?;
    if observed.as_ref() != before {
        return Err(Error::PatchConflict);
    }
    let metadata_source = pop_up_menu_metadata::strict_source(source)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    budget.charge_wire_bytes(metadata_source.payload.len(), path)?;
    let metadata_options = popup_metadata_options(metadata_source.payload.len(), 1, &budget);
    let metadata_facts = pop_up_menu_metadata::inspect(source, metadata_options)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let metadata_report = metadata_facts.report();
    budget.charge_wire_fields(metadata_report.fields(), path)?;
    budget.charge_wire_work(metadata_report.work_bytes(), path)?;
    budget.charge_wire_nesting(metadata_report.max_depth(), path)?;
    budget.charge_payload_items(metadata_report.components_scanned(), path)?;
    budget.charge_payload_references(metadata_report.references_scanned(), path)?;
    budget.charge_transaction_work(
        metadata_report
            .input_bytes()
            .saturating_add(metadata_report.output_bytes()),
        path,
    )?;
    if metadata_facts.has_physical_alias() {
        return Err(Error::Refused(Refusal::UnsupportedDependency { path }));
    }
    metadata_facts
        .require_single_current_component(core::slice::from_ref(&target.component_index))
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let fresh = metadata_facts
        .allocate_identifiers(1)
        .map_err(|error| map_popup_metadata_error(error, path))?
        .into_iter()
        .next()
        .ok_or(Error::InvalidSource { path })?;
    let native = rewrite_native_entry(
        source,
        target,
        after,
        fresh.identifier,
        path,
        &metadata_facts,
        &mut budget,
    )?;
    let native_component_indices = native
        .member_edits
        .iter()
        .map(|edit| edit.component_index)
        .collect::<Vec<_>>();
    let touched_component_indices = metadata_facts
        .dedupe_touched_components(&native_component_indices)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    // Resolve the complete current/effective selector set before staging the
    // metadata candidate. This is also the source of truth for the one-token
    // transition across every changed native member.
    metadata_facts
        .touched_selectors(&touched_component_indices)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let metadata_payload = rewrite_popup_metadata(
        source,
        &metadata_facts,
        &touched_component_indices,
        fresh,
        native.added_model,
        native.added_model_component,
        native.control_component_index,
        &native.removed_models,
        &native.removed_model_components,
        register_model_uuid,
        path,
        &mut budget,
    )?;
    let metadata_bytes = pack_metadata_payload(
        source,
        metadata_facts.route(),
        metadata_payload,
        path,
        &mut budget,
    )?;
    let previews = if delete_previews {
        crate::structure::root_previews(catalog).map_err(|_| Error::InvalidSource { path })?
    } else {
        Vec::new()
    };
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(native.member_edits.len().saturating_add(1))
        .map_err(|_| Error::Allocation {
            amount: native.member_edits.len().saturating_add(1),
            path,
        })?;
    for edit in &native.member_edits {
        edits.push(EntryEdit::new(
            edit.member_name.as_str(),
            edit.member_bytes.as_slice(),
        ));
    }
    edits.push(EntryEdit::new(METADATA_ENTRY, &metadata_bytes));
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &previews, catalog.limits())
        .map_err(|_| Error::Verification)?;
    let requirements = prepared.execution_requirements();
    // Candidate reassembly and the subsequent full Package reopen are both
    // precharged before the final ZIP output allocation.  This owner retains a
    // staged-private policy: native/metadata candidates are private until the
    // ZIP candidate and semantic readback succeed.
    budget.preflight_reassembly(requirements, path)?;
    budget.charge_transaction_work(
        requirements
            .output_bytes()
            .saturating_mul(2)
            .saturating_add(source.catalog().len().saturating_mul(1024)),
        path,
    )?;
    budget.charge_candidate_input_bytes(requirements.output_bytes(), path)?;
    let bytes = prepared
        .execute(requirements.exact_limits())
        .map_err(|_| Error::Verification)?;
    let candidate = reopen(bytes)?;
    let candidate_catalog = candidate.physical().ok_or(Error::Verification)?;
    budget.charge_candidate_reopen(candidate_catalog, path)?;
    let native_member_names = native
        .member_edits
        .iter()
        .map(|edit| edit.member_name.as_str())
        .collect::<Vec<_>>();
    verify_locality(catalog, candidate_catalog, &native_member_names, path)?;
    let candidate_target = retarget(&candidate)?;
    let reread = read_popup_with_budget(&candidate, candidate_target, &mut budget, true)?;
    if reread.as_ref() != after {
        return Err(Error::Verification);
    }
    let mut diagnostic_components = touched_component_indices.clone();
    if !diagnostic_components.contains(&metadata_facts.route().component_index) {
        diagnostic_components.push(metadata_facts.route().component_index);
    }
    Ok(Publication {
        package: candidate,
        touched_components: diagnostic_components.len(),
        deleted_previews: previews.len(),
    })
}

/// Verify package-level locality for a native transaction that rewrites one
/// or more physical members. All non-authorized members must remain exact,
/// including opaque/versioned entries and metadata; root previews may only
/// be deleted.
///
/// # Errors
///
/// [`Error::Verification`] for any other change.
pub fn verify_locality(
    source_catalog: &SourceCatalog,
    candidate_catalog: &SourceCatalog,
    native_members: &[&str],
    _path: Path,
) -> Result<(), Error> {
    if native_members.iter().enumerate().any(|(index, name)| {
        name.is_empty()
            || *name == METADATA_ENTRY
            || name.starts_with("preview")
            || native_members[..index].contains(name)
    }) {
        return Err(Error::Verification);
    }
    let selected = |name: &str| native_members.contains(&name) || name == METADATA_ENTRY;

    for name in native_members
        .iter()
        .copied()
        .chain(std::iter::once(METADATA_ENTRY))
    {
        let mut source_matches = source_catalog
            .package()
            .iter()
            .filter(|entry| entry.name() == name);
        let source_entry = source_matches.next().ok_or(Error::Verification)?;
        if source_matches.next().is_some() {
            return Err(Error::Verification);
        }
        let mut candidate_matches = candidate_catalog
            .package()
            .iter()
            .filter(|entry| entry.name() == name);
        let candidate_entry = candidate_matches.next().ok_or(Error::Verification)?;
        if candidate_matches.next().is_some()
            || !super::locality::selected_package_member_preserved(source_entry, candidate_entry)
        {
            return Err(Error::Verification);
        }
    }

    // Preview invalidation may delete source previews, but it may neither add
    // one nor rewrite a retained preview's physical ZIP record.
    for candidate_entry in candidate_catalog
        .package()
        .iter()
        .filter(|entry| entry.name().starts_with("preview"))
    {
        let mut source_matches = source_catalog
            .package()
            .iter()
            .filter(|entry| entry.name() == candidate_entry.name());
        let source_entry = source_matches.next().ok_or(Error::Verification)?;
        if source_matches.next().is_some()
            || !super::locality::package_member_preserved(source_entry, candidate_entry)
        {
            return Err(Error::Verification);
        }
    }

    let mut source_untouched = source_catalog
        .package()
        .iter()
        .filter(|entry| !selected(entry.name()) && !entry.name().starts_with("preview"));
    let mut candidate_untouched = candidate_catalog
        .package()
        .iter()
        .filter(|entry| !selected(entry.name()) && !entry.name().starts_with("preview"));
    loop {
        match (source_untouched.next(), candidate_untouched.next()) {
            (Some(source_entry), Some(candidate_entry))
                if source_entry.name() == candidate_entry.name()
                    && super::locality::package_member_preserved(source_entry, candidate_entry) => {
            },
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
    }
    Ok(())
}
