//! A populated copy of a body table (record 0371; the migration host's
//! `duplicate_table`).
//!
//! The copy's graph — fresh identifiers for its drawable, model and owned
//! storage, the host's table UUID, its CalculationEngine owner family and
//! formula-list rewrites, and the clones' registry records — is the shared
//! engine `litchi_iwa_table::structure::duplicate` (record 0370h). This
//! module attaches the copy to the body as the host does: a clone of the
//! source table's drawable attachment naming the new drawable, one
//! object-replacement character and attachment-table entry at the requested
//! position, and the references of the body, the engine and the attachment
//! to the copy across components.

use litchi_iwa_core::RawMessage;
use litchi_iwa_table::structure::duplicate::{self as engine, Drawable};
use litchi_iwa_table::structure::metadata::{self, Transition};
use litchi_iwa_table::structure::{self, Fault, FaultLimit, Working, references};
use litchi_numbers_wire::table_create;

use super::Package;
use super::body_table_creation::{
    Commit, Diagnostics, Error, LimitKind, Patch, Refusal, anchored_body, prepared_len,
};
use super::structure::Artifacts;
use super::table_lock::{self, BodyTableTarget, OwnedBodyTableSelector};

/// `TST.TableInfoArchive` and its legacy form.
const TABLE_INFO_MESSAGE_TYPES: [u32; 2] = [6_000, 6_003];
/// `TSWP.StorageArchive` (the body story).
const BODY_STORAGE_TYPES: [u32; 2] = [2_001, 2_022];
/// `TSD.DrawableAttachmentArchive.drawable`.
const ATTACHMENT_DRAWABLE_FIELD: u32 = 1;
/// How far an inline copy is offset from its source (the host's
/// `INLINE_TABLE_DUPLICATE_OFFSET`).
const OFFSET: f32 = 0.0;

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Unsupported => Self::Refused(Refusal::Unsupported),
            Fault::Dependency => Self::Refused(Refusal::UnsupportedDependency),
            Fault::Invalid => Self::InvalidSource,
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Verification => Self::Verification,
        }
    }
}

/// A body-table discovery failure: the selector's refusals stay typed; a
/// body or table outside the rooted graph this transaction reads is
/// [`Refusal::Unsupported`], as for creation's template.
pub(super) fn lock_error(error: table_lock::Error) -> Error {
    use crate::table::lock::Refusal as Lock;
    use table_lock::LimitKind as Kind;
    match error {
        table_lock::Error::Refused(Lock::TableNotFound) => Error::Refused(Refusal::TableNotFound),
        table_lock::Error::Refused(Lock::AmbiguousTableName) => {
            Error::Refused(Refusal::AmbiguousTableName)
        },
        table_lock::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                Kind::InputBytes
                | Kind::OutputBytes
                | Kind::Entries
                | Kind::EntryBytes
                | Kind::TotalEntryBytes
                | Kind::PackageBytes => LimitKind::PackageBytes,
                _ => LimitKind::PayloadBytes,
            },
            observed,
            maximum,
        },
        table_lock::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::Refused(Refusal::Unsupported),
    }
}

/// The source table `selector` names among the rooted body tables.
fn select<'t>(
    targets: &'t [BodyTableTarget],
    selector: &OwnedBodyTableSelector,
) -> Result<&'t BodyTableTarget, Error> {
    match selector {
        OwnedBodyTableSelector::Position(position) => targets
            .get(position.get())
            .ok_or(Error::Refused(Refusal::TableNotFound)),
        OwnedBodyTableSelector::Name(name) => {
            let mut matches = targets
                .iter()
                .filter(|target| target.table_name.as_ref() == name.as_ref());
            let first = matches
                .next()
                .ok_or(Error::Refused(Refusal::TableNotFound))?;
            if matches.next().is_some() {
                return Err(Error::Refused(Refusal::AmbiguousTableName));
            }
            Ok(first)
        },
    }
}

/// Copy the body table `selector` names and anchor the copy at UTF-16
/// position `anchor` of the body story.
pub(super) fn duplicate(
    source: &Package,
    selector: &OwnedBodyTableSelector,
    anchor: usize,
) -> Result<Commit, Error> {
    let mut budget =
        table_lock::WireBudget::new(source.state.source.limits()).map_err(lock_error)?;
    let targets = table_lock::native_body_table_targets_with_budget(source, &mut budget)
        .map_err(lock_error)?;
    let target = select(&targets, selector)?;
    let names: Vec<&str> = targets
        .iter()
        .map(|target| target.table_name.as_ref())
        .collect();
    let name = engine::copy_name(&target.table_name, &names).ok_or(Error::InvalidSource)?;
    let info = target.drawable_identifier.get();
    let template_attachment = target.attachment_identifier.get();

    let mut working = Working::new(source)?;
    working.keep_previews();
    let first = working.next_identifier();
    let mut transition = Transition::default();
    let copy = engine::clone_graph(
        &mut working,
        &mut transition,
        info,
        Drawable {
            message_types: &TABLE_INFO_MESSAGE_TYPES,
            info_path: &[],
            offset: OFFSET,
            parent: None,
        },
        &name,
    )?;
    let new_info = copy.info;
    let info_object = working.object(new_info)?;
    let info_index = structure::unique_message(info_object, &TABLE_INFO_MESSAGE_TYPES)?;
    let body = table_create::drawable_parent(&info_object.messages[info_index].data)
        .map_err(|_error| Error::InvalidSource)?
        .ok_or(Error::Refused(Refusal::Unsupported))?;

    // The attachment: a clone of the source's naming the new drawable.
    let attachment = *working.allocate(1)?.first().ok_or(Error::InvalidSource)?;
    let attachment_component = working
        .locate(template_attachment)?
        .ok_or(Error::InvalidSource)?;
    let limits = working.archive_limits();
    let template = working.object(template_attachment)?;
    let [template_message] = template.messages.as_slice() else {
        return Err(Error::Refused(Refusal::Unsupported));
    };
    if table_create::single_reference(&template_message.data, ATTACHMENT_DRAWABLE_FIELD)
        .map_err(|_error| Error::InvalidSource)?
        != Some(info)
    {
        return Err(Error::Refused(Refusal::Unsupported));
    }
    let data = litchi_iwa_common::wire::patch_length_delimited_field(
        &template_message.data,
        ATTACHMENT_DRAWABLE_FIELD,
        true,
        Some(&table_create::reference_payload(new_info)),
    )
    .map_err(|_error| Error::InvalidSource)?;
    let object = template
        .clone_with_identity_remap(
            attachment,
            &[(template_attachment, attachment), (info, new_info)],
            &[RawMessage {
                type_: template_message.type_,
                data,
            }],
            limits,
        )
        .map_err(|error| Error::from(structure::map_core(&error)))?;
    working.insert(&attachment_component, object)?;

    // The body story: the anchor character and its attachment entry.
    let body_component = working.locate(body)?.ok_or(Error::InvalidSource)?;
    let body_object = working.object(body)?;
    let body_index = structure::unique_message(body_object, &BODY_STORAGE_TYPES)?;
    let body_payload = &body_object.messages[body_index].data;
    if anchor > prepared_len(body_payload)? {
        return Err(Error::Refused(Refusal::PositionOutOfBounds));
    }
    let body_data = anchored_body(source, body_payload, anchor, attachment)?;
    references::replace(
        working.object_mut(body)?,
        body_index,
        body_data,
        &[],
        &[attachment],
        &[],
        limits,
    )?;

    // Registry: the attachment's UUID (when its template has one), and the
    // body's, the engine's and the attachment's references to the copy, in
    // the host's order.
    if let Some(registry) = metadata::registry(&working)? {
        let component = |name: &str| -> Result<u64, Error> {
            Ok(registry.component_for_entry(name)?.identifier)
        };
        let info_component = component(&copy.info_component)?;
        let attachment_id = component(&attachment_component)?;
        transition.reference(
            &registry,
            component(&body_component)?,
            info_component,
            new_info,
        )?;
        if let Some(engine_component) = &copy.engine_component {
            transition.reference(
                &registry,
                component(engine_component)?,
                info_component,
                new_info,
            )?;
        }
        if registry
            .uuids
            .iter()
            .any(|binding| binding.object == template_attachment)
        {
            let seed = metadata::identity_seed(&working, attachment)?;
            transition.register_uuid(&registry, attachment_id, attachment, seed)?;
        }
        transition.reference(&registry, attachment_id, info_component, new_info)?;
        transition.reference(
            &registry,
            component(&body_component)?,
            attachment_id,
            attachment,
        )?;
    }
    let created_objects = usize::try_from(working.next_identifier() - first)
        .map_err(|_error| Error::InvalidSource)?;
    let published = working.publish(&transition)?;
    let package = published.package;

    // The copy reads back at the anchor with its name and size.
    let mut budget = table_lock::WireBudget::new(package.state.source.limits())
        .map_err(|_error| Error::Verification)?;
    let after = table_lock::native_body_table_targets_with_budget(&package, &mut budget)
        .map_err(|_error| Error::Verification)?;
    let position = after
        .iter()
        .position(|candidate| candidate.drawable_identifier.get() == new_info)
        .ok_or(Error::Verification)?;
    let created = &after[position];
    if after.len() != targets.len() + 1
        || created.table_name.as_ref() != name
        || created.anchor_character_index != anchor
        || (created.table_rows, created.table_columns) != (target.table_rows, target.table_columns)
    {
        return Err(Error::Verification);
    }
    Ok(Commit {
        patch: Patch {
            artifacts: Artifacts::new(source, &package),
            table: position,
        },
        package,
        diagnostics: Diagnostics {
            changed: true,
            created_objects,
            touched_components: published.touched_components,
            seeded_cells: 0,
        },
    })
}
