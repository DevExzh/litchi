//! Cached focused handoff for Pages body-table merge reads.
//!
//! The compatibility editor receives native model identifiers for historical
//! reasons.  This adapter keeps that identity below the public focused reader
//! boundary: it proves the body-owned table order from the parsed archive
//! catalog, converts the selected model to a positional selector, and only
//! then admits the shared focused merge reader. Source-built packages retain
//! the editor's historical focused compatibility path; an exact package with
//! no rooted owner returns the focused reader's typed not-found result.

use std::collections::{HashMap, HashSet};
use std::num::NonZeroU64;

use litchi_iwa_archive::ComponentCatalog;
use litchi_iwa_common::wire::{WireFieldView, WireView};
use litchi_iwa_common::{WireLimits, decode_varint_from_bytes};
use litchi_iwa_protos::table_info_codec::{self, WireResourceLimit};
use litchi_numbers::table::merge::Region;
use litchi_pages::{
    BodyTableMergesError, BodyTableMergesLimitKind, BodyTableSelector, MergeReader,
};

use super::super::PagesEditor;
use crate::archive::{ArchiveObject, RawMessage};
use crate::{Error, IWorkPackage, Result};

const BODY_STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const TABLE_ATTACHMENT_FIELD: u32 = 9;
const TABLE_ENTRY_FIELD: u32 = 1;
const TABLE_ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const TABLE_ENTRY_OBJECT_FIELD: u32 = 2;
const DRAWABLE_ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const DRAWABLE_FIELD: u32 = 1;
const TABLE_INFO_MESSAGE_TYPES: &[u32] = &[6_000, 6_003];
const MAX_BODY_TABLES: usize = 4_096;

/// Read merged-cell regions through the cached focused Pages owner.
pub(super) fn regions_in_editor(editor: &PagesEditor, model_object_id: u64) -> Result<Vec<Region>> {
    if editor.package().exact_source_owner().is_none() {
        let (focused, table_position, _) =
            super::semantic::focused_body_table_source(editor, model_object_id)?;
        return focused
            .body_table_merges(BodyTableSelector::index(table_position))
            .map_err(map_focused_merges_error);
    }

    let (components, limits) = editor
        .package()
        .shared_component_catalog(map_catalog_error)?;
    let Some(table_position) = focused_table_position(
        editor.package(),
        components.as_ref(),
        editor.body_storage_id.get(),
        model_object_id,
    )?
    else {
        return Err(Error::ParseError(format!(
            "Pages table model {model_object_id} is not attached to the body"
        )));
    };
    let reader =
        MergeReader::__from_shared_catalog(components, limits).map_err(map_focused_merges_error)?;
    reader
        .body_table_merges(BodyTableSelector::index(table_position))
        .map_err(map_focused_merges_error)
}

/// Resolve a compatibility model identifier to the position used by the
/// focused Pages reader.  This scan projects only storage table entries,
/// attachment references, and `TableInfoArchive` ownership; in particular it
/// never decodes a `TableModelArchive` or materializes body text.
fn focused_table_position(
    package: &IWorkPackage,
    components: &ComponentCatalog,
    body_storage_id: u64,
    model_object_id: u64,
) -> Result<Option<usize>> {
    let mut budget = SelectorBudget::new(components)?;
    let locations = ObjectLocations::build(components, &mut budget)?;
    let body = locations.object(body_storage_id)?;
    let body_message = unique_message(body, BODY_STORAGE_MESSAGE_TYPES, "Pages body storage")?;
    let entries = decode_table_entries(package, body_message.data.as_slice(), &mut budget)?;
    let mut seen_models = HashSet::new();
    seen_models
        .try_reserve(entries.len())
        .map_err(|_| allocation_error(entries.len()))?;

    let mut selected = None;
    for (position, entry) in entries.iter().enumerate() {
        budget.charge_references(2)?;
        let attachment = locations.object(entry.attachment_id.get())?;
        let attachment_message = unique_message(
            attachment,
            &[DRAWABLE_ATTACHMENT_MESSAGE_TYPE],
            "Pages drawable attachment",
        )?;
        let drawable_id = parse_drawable_id(package, attachment_message.data.as_slice())?;
        budget.charge_work(attachment_message.data.len().saturating_mul(8))?;
        let drawable = locations.object(drawable_id.get())?;
        let table_info = unique_message(drawable, TABLE_INFO_MESSAGE_TYPES, "Pages table info")?;
        budget.charge_work(table_info.data.len().saturating_mul(8))?;
        let table_info = decode_table_info(package, table_info.data.as_slice(), body_storage_id)?;
        let model_id = table_info.table_model().identifier().get();
        if !seen_models.insert(model_id) {
            return Err(Error::InvalidFormat(format!(
                "Pages body table model {model_id} is attached more than once"
            )));
        }
        if model_id == model_object_id && selected.replace(position).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Pages body table model {model_object_id} has multiple rooted owners"
            )));
        }
    }
    Ok(selected)
}

#[derive(Debug, Clone, Copy)]
struct TableEntry {
    character_index: u32,
    attachment_id: NonZeroU64,
}

fn decode_table_entries(
    package: &IWorkPackage,
    source: &[u8],
    budget: &mut SelectorBudget,
) -> Result<Vec<TableEntry>> {
    budget.charge_work(source.len().saturating_mul(8))?;
    let limits = wire_limits(package, source)?;
    let view = WireView::parse_with_limits(source, limits)
        .map_err(|error| invalid_wire("Pages body storage", error))?;
    let table = unique_field(&view, TABLE_ATTACHMENT_FIELD, 2)?;
    let Some(table) = table else {
        return Ok(Vec::new());
    };
    let table_view = WireView::parse_with_limits(table.payload(), limits)
        .map_err(|error| invalid_wire("Pages body table attachments", error))?;
    let mut entries = Vec::new();
    for field in table_view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRY_FIELD)
    {
        budget.charge_items(1)?;
        budget.charge_work(field.raw().len().saturating_mul(8))?;
        validate_field(field, 2)?;
        let entry = WireView::parse_with_limits(field.payload(), limits)
            .map_err(|error| invalid_wire("Pages body table entry", error))?;
        let character_index = unique_field(&entry, TABLE_ENTRY_CHARACTER_INDEX_FIELD, 0)?
            .ok_or_else(|| invalid_source("Pages body table entry has no character index"))?;
        let character_index = parse_u32(character_index)?;
        let attachment = unique_field(&entry, TABLE_ENTRY_OBJECT_FIELD, 2)?
            .ok_or_else(|| invalid_source("Pages body table entry has no attachment"))?;
        budget.charge_references(1)?;
        let attachment_id = parse_local_reference(attachment.payload(), limits)?;
        entries.try_reserve(1).map_err(|_| allocation_error(1))?;
        entries.push(TableEntry {
            character_index,
            attachment_id,
        });
    }
    entries.sort_unstable_by_key(|entry| entry.character_index);
    if entries
        .windows(2)
        .any(|pair| pair[0].character_index == pair[1].character_index)
    {
        return Err(invalid_source(
            "Pages body table entries repeat an anchor position",
        ));
    }
    Ok(entries)
}

fn parse_drawable_id(package: &IWorkPackage, source: &[u8]) -> Result<NonZeroU64> {
    let limits = wire_limits(package, source)?;
    let view = WireView::parse_with_limits(source, limits)
        .map_err(|error| invalid_wire("Pages drawable attachment", error))?;
    let drawable = unique_field(&view, DRAWABLE_FIELD, 2)?
        .ok_or_else(|| invalid_source("Pages drawable attachment has no drawable"))?;
    parse_local_reference(drawable.payload(), limits)
}

fn parse_local_reference(source: &[u8], limits: WireLimits) -> Result<NonZeroU64> {
    let view = WireView::parse_with_limits(source, limits)
        .map_err(|error| invalid_wire("Pages local reference", error))?;
    let mut identifier = None;
    let mut external = None;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|error| invalid_wire("Pages local reference", error))?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(invalid_source(
                        "Pages local reference identifier is invalid",
                    ));
                }
                let (value, length) = decode_varint_from_bytes(field.payload())
                    .map_err(|error| invalid_wire("Pages local reference identifier", error))?;
                if length != field.payload().len() {
                    return Err(invalid_source(
                        "Pages local reference identifier is noncanonical",
                    ));
                }
                identifier = NonZeroU64::new(value);
                if identifier.is_none() {
                    return Err(invalid_source("Pages local reference identifier is zero"));
                }
            },
            3 => {
                if field.wire_type() != 0 || external.is_some() {
                    return Err(invalid_source(
                        "Pages local reference external flag is invalid",
                    ));
                }
                let (value, length) = decode_varint_from_bytes(field.payload())
                    .map_err(|error| invalid_wire("Pages local reference external flag", error))?;
                if length != field.payload().len() || value > 1 {
                    return Err(invalid_source(
                        "Pages local reference external flag is noncanonical",
                    ));
                }
                external = Some(value != 0);
            },
            _ => {},
        }
    }
    if external == Some(true) {
        return Err(invalid_source("Pages table reference is external"));
    }
    identifier.ok_or_else(|| invalid_source("Pages local reference has no identifier"))
}

fn decode_table_info(
    package: &IWorkPackage,
    source: &[u8],
    body_storage_id: u64,
) -> Result<table_info_codec::TableInfoSnapshot> {
    let options = table_info_options(package, source)?;
    let snapshot = table_info_codec::decode_table_info_with_parent(source, options)
        .map_err(|error| map_table_info_error(body_storage_id, error))?;
    if snapshot.parent() != NonZeroU64::new(body_storage_id) {
        return Err(invalid_source("Pages table info is not owned by the body"));
    }
    Ok(snapshot)
}

fn table_info_options(
    package: &IWorkPackage,
    source: &[u8],
) -> Result<table_info_codec::DecodeOptions> {
    let limits = wire_limits(package, source)?;
    let recursion_limit = u32::try_from(limits.max_nesting())
        .map_err(|_| invalid_source("Pages table-info nesting limit overflows u32"))?;
    Ok(table_info_codec::DecodeOptions::new(
        limits.max_input_bytes(),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion_limit,
    ))
}

fn wire_limits(package: &IWorkPackage, source: &[u8]) -> Result<WireLimits> {
    let archive_limits = package.limits().archive_limits();
    let input = source
        .len()
        .min(archive_limits.max_message_bytes())
        .min(archive_limits.max_archive_bytes())
        .min(package.limits().max_iwa_stream_bytes())
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    WireLimits::default()
        .with_input_bytes(input)
        .and_then(|limits| {
            limits.with_fields(
                source
                    .len()
                    .saturating_mul(2)
                    .clamp(1, WireLimits::MAX_FIELDS),
            )
        })
        .and_then(|limits| {
            limits.with_rewrite_work(
                source
                    .len()
                    .saturating_mul(8)
                    .clamp(1, WireLimits::MAX_REWRITE_WORK),
            )
        })
        .map_err(|error| {
            Error::InvalidFormat(format!("invalid Pages merge selector limits: {error}"))
        })
}

fn unique_field<'source>(
    view: &'source WireView<'source>,
    number: u32,
    wire_type: u8,
) -> Result<Option<WireFieldView<'source>>> {
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == number) {
        validate_field(field, wire_type)?;
        if selected.replace(field).is_some() {
            return Err(invalid_source(
                "Pages merge selector repeats a singular field",
            ));
        }
    }
    Ok(selected)
}

fn validate_field(field: WireFieldView<'_>, wire_type: u8) -> Result<()> {
    if field.wire_type() != wire_type {
        return Err(invalid_source(
            "Pages merge selector field has the wrong wire type",
        ));
    }
    field
        .validate_canonical_framing()
        .map_err(|error| invalid_wire("Pages merge selector field", error))
}

fn parse_u32(field: WireFieldView<'_>) -> Result<u32> {
    let (value, length) = decode_varint_from_bytes(field.payload())
        .map_err(|error| invalid_wire("Pages table entry character index", error))?;
    if length != field.payload().len() {
        return Err(invalid_source(
            "Pages table entry character index is noncanonical",
        ));
    }
    u32::try_from(value)
        .map_err(|_| invalid_source("Pages table entry character index exceeds u32"))
}

fn unique_message<'object>(
    object: &'object ArchiveObject,
    types: &[u32],
    name: &str,
) -> Result<&'object RawMessage> {
    let mut selected = None;
    for message in &object.messages {
        if !types.contains(&message.type_) {
            continue;
        }
        if selected.replace(message).is_some() {
            return Err(invalid_source(format!("{name} has duplicate payloads")));
        }
    }
    selected.ok_or_else(|| invalid_source(format!("{name} payload is missing")))
}

struct ObjectLocations<'catalog> {
    objects: HashMap<u64, &'catalog ArchiveObject>,
}

/// Aggregate admission ledger for the host-side identity projection.
///
/// Each nested wire view has its own local profile, but the selector must not
/// multiply that profile by the number of body entries. This ledger is built
/// before any selector collection is reserved and charges logical entries,
/// references, and projected wire work across the complete catalog.
struct SelectorBudget {
    objects: usize,
    max_objects: usize,
    items: usize,
    max_items: usize,
    references: usize,
    max_references: usize,
    work: usize,
    max_work: usize,
}

impl SelectorBudget {
    fn new(components: &ComponentCatalog) -> Result<Self> {
        let mut object_count = 0usize;
        let mut payload_bytes = 0usize;
        for component in components.iter() {
            object_count = object_count
                .checked_add(component.archive().objects.len())
                .ok_or_else(|| allocation_error(usize::MAX))?;
            for object in &component.archive().objects {
                for message in &object.messages {
                    payload_bytes = payload_bytes
                        .checked_add(message.data.len())
                        .ok_or_else(|| allocation_error(usize::MAX))?;
                }
            }
        }
        let max_references = object_count
            .saturating_mul(4)
            .clamp(1, WireLimits::MAX_FIELDS);
        let max_work = payload_bytes
            .saturating_mul(16)
            .clamp(1, WireLimits::MAX_REWRITE_WORK);
        Ok(Self {
            objects: 0,
            max_objects: WireLimits::MAX_FIELDS,
            items: 0,
            // Logical body entries are distinct from ZIP entries. Keep the
            // format's focused table ceiling so one component can own many
            // attachments while still bounding collection growth.
            max_items: MAX_BODY_TABLES,
            references: 0,
            max_references,
            work: 0,
            max_work,
        })
    }

    fn charge_objects(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.objects,
            amount,
            self.max_objects,
            BodyTableMergesLimitKind::PayloadObjects,
        )
    }

    fn charge_items(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.items,
            amount,
            self.max_items,
            BodyTableMergesLimitKind::PayloadItems,
        )
    }

    fn charge_references(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.references,
            amount,
            self.max_references,
            BodyTableMergesLimitKind::PayloadReferences,
        )
    }

    fn charge_work(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.work,
            amount,
            self.max_work,
            BodyTableMergesLimitKind::WireWork,
        )
    }

    fn charge(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: BodyTableMergesLimitKind,
    ) -> Result<()> {
        let observed = current.saturating_add(amount);
        if observed > maximum {
            return Err(Error::PagesTableMerges(
                BodyTableMergesError::LimitExceeded {
                    kind,
                    observed: observed as u64,
                    maximum: maximum as u64,
                },
            ));
        }
        *current = observed;
        Ok(())
    }
}

impl<'catalog> ObjectLocations<'catalog> {
    fn build(components: &'catalog ComponentCatalog, budget: &mut SelectorBudget) -> Result<Self> {
        let object_count = components
            .iter()
            .try_fold(0usize, |count, component| {
                count.checked_add(component.archive().objects.len())
            })
            .ok_or_else(|| allocation_error(usize::MAX))?;
        budget.charge_objects(object_count)?;
        let mut objects = HashMap::new();
        objects
            .try_reserve(object_count)
            .map_err(|_| allocation_error(object_count))?;
        for component in components.iter() {
            for object in &component.archive().objects {
                let Some(identifier) = object.archive_info.identifier else {
                    continue;
                };
                if objects.insert(identifier, object).is_some() {
                    return Err(invalid_source(format!(
                        "Pages object {identifier} occurs in multiple components"
                    )));
                }
            }
        }
        Ok(Self { objects })
    }

    fn object(&self, identifier: u64) -> Result<&'catalog ArchiveObject> {
        self.objects
            .get(&identifier)
            .copied()
            .ok_or_else(|| invalid_source(format!("Pages object {identifier} is missing")))
    }
}

fn map_focused_merges_error(error: BodyTableMergesError) -> Error {
    Error::PagesTableMerges(error)
}

fn map_catalog_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Iwa(error) => error.into(),
        litchi_iwa_archive::Error::Allocation { resource, amount } => {
            litchi_iwa_common::Error::Allocation { resource, amount }.into()
        },
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableMergesError::LimitExceeded {
            kind: map_archive_limit(kind),
            observed,
            maximum,
        }
        .into(),
        error => error.into(),
    }
}

fn map_archive_limit(kind: litchi_iwa_archive::LimitKind) -> BodyTableMergesLimitKind {
    use litchi_iwa_archive::LimitKind as ArchiveLimitKind;
    match kind {
        ArchiveLimitKind::InputBytes => BodyTableMergesLimitKind::InputBytes,
        ArchiveLimitKind::OutputBytes => BodyTableMergesLimitKind::OutputBytes,
        ArchiveLimitKind::Entries => BodyTableMergesLimitKind::Entries,
        ArchiveLimitKind::MemberNameBytes => BodyTableMergesLimitKind::EntryBytes,
        ArchiveLimitKind::MetadataBytes => BodyTableMergesLimitKind::PackageBytes,
        ArchiveLimitKind::CompressedEntryBytes | ArchiveLimitKind::EntryBytes => {
            BodyTableMergesLimitKind::EntryBytes
        },
        ArchiveLimitKind::TotalBytes => BodyTableMergesLimitKind::TotalEntryBytes,
        ArchiveLimitKind::IwaStreamBytes => BodyTableMergesLimitKind::PayloadBytes,
        ArchiveLimitKind::IwaTotalBytes => BodyTableMergesLimitKind::TotalPayloadBytes,
    }
}

fn map_table_info_error(body_storage_id: u64, error: table_info_codec::DecodeError) -> Error {
    let limit = if let Some((observed, maximum)) = error.field_limit_values() {
        Some((BodyTableMergesLimitKind::WireFields, observed, maximum))
    } else if let Some((observed, maximum)) = error.work_limit_values() {
        Some((BodyTableMergesLimitKind::WireWork, observed, maximum))
    } else {
        match error.wire_resource_limit() {
            Some(WireResourceLimit::Bytes { observed, maximum }) => Some((
                BodyTableMergesLimitKind::WireBytes,
                observed.unwrap_or(usize::MAX),
                maximum.unwrap_or(usize::MAX),
            )),
            Some(WireResourceLimit::Nesting { observed, maximum }) => Some((
                BodyTableMergesLimitKind::WireNesting,
                observed
                    .and_then(|value| usize::try_from(value).ok())
                    .unwrap_or(usize::MAX),
                usize::try_from(maximum.unwrap_or(u32::MAX)).unwrap_or(usize::MAX),
            )),
            Some(_) => None,
            None => None,
        }
    };
    if let Some((kind, observed, maximum)) = limit {
        return BodyTableMergesError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        }
        .into();
    }
    Error::InvalidFormat(format!(
        "Pages table-info ownership for body {body_storage_id} failed strict validation: {error}"
    ))
}

fn allocation_error(amount: usize) -> Error {
    BodyTableMergesError::Allocation { amount }.into()
}

fn invalid_source(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn invalid_wire(context: &str, error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!("{context} wire projection failed: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn body_selector_budget_rejects_4097_logical_attachments() {
        let mut table = Vec::new();
        for character_index in 0..=MAX_BODY_TABLES {
            let mut entry = Vec::new();
            litchi_iwa_common::wire::append_varint_field(
                &mut entry,
                TABLE_ENTRY_CHARACTER_INDEX_FIELD,
                character_index as u64,
            )
            .expect("test character index fits the wire format");
            let mut attachment = Vec::new();
            litchi_iwa_common::wire::append_varint_field(&mut attachment, 1, 1)
                .expect("test attachment identifier fits the wire format");
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut entry,
                TABLE_ENTRY_OBJECT_FIELD,
                &attachment,
            )
            .expect("test attachment reference fits the wire format");
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut table,
                TABLE_ENTRY_FIELD,
                &entry,
            )
            .expect("test table entry fits the wire format");
        }
        let mut body = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut body,
            TABLE_ATTACHMENT_FIELD,
            &table,
        )
        .expect("test table attachment payload fits the wire format");

        let components = ComponentCatalog::__empty();
        let mut budget = SelectorBudget::new(&components).expect("empty catalog has a budget");
        budget.max_references = WireLimits::MAX_FIELDS;
        budget.max_work = WireLimits::MAX_REWRITE_WORK;
        let error = decode_table_entries(&IWorkPackage::new(), &body, &mut budget)
            .expect_err("the focused table ceiling must reject the 4097th attachment");
        assert!(matches!(
            error,
            Error::PagesTableMerges(BodyTableMergesError::LimitExceeded {
                kind: BodyTableMergesLimitKind::PayloadItems,
                observed: 4_097,
                maximum: 4_096,
            })
        ));
    }
}
