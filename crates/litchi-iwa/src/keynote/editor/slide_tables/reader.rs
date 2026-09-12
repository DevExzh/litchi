//! Cached focused handoff for Keynote slide-table merge reads.
//!
//! The compatibility editor accepts native table model identifiers for
//! historical reasons. This adapter proves the rooted slide/table order from
//! the parsed archive catalog, converts that identity to a positional selector,
//! and then admits the focused merge reader. The selector projection retains
//! only object references and table-info facts; table-model payloads remain
//! owned by the focused reader.

use std::collections::{HashMap, HashSet};
use std::num::NonZeroU64;

use litchi_iwa_archive::{ComponentCatalog, Error as ArchiveError, LimitKind as ArchiveLimitKind};
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_protos::{
    keynote_document_codec, keynote_show_codec,
    table_info_codec::{self, WireResourceLimit},
};
use litchi_keynote::{MergeReader, SlideSelector, SlideTableMergesError, TableSelector};
use litchi_numbers::table::merge::Region;

use super::super::KeynoteEditor;
use crate::archive::{ArchiveObject, RawMessage};
use crate::{Error, IWorkPackage, Result};

const ROOT_OBJECT_ID: u64 = 1;
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHOW_MESSAGE_TYPE: u32 = 2;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE: u32 = 6_200;
const COLUMN_ROW_UID_MAP_MESSAGE_TYPE: u32 = 6_267;
const OWNED_DRAWABLE_FIELD: u32 = 7;
const DRAWABLE_Z_ORDER_FIELD: u32 = 42;
const ROLE_MESSAGE_TYPES: &[u32] = &[
    TABLE_INFO_MESSAGE_TYPE,
    TABLE_MODEL_MESSAGE_TYPE,
    TILE_MESSAGE_TYPE,
    HEADER_BUCKET_MESSAGE_TYPE,
    COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE,
    COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
    TABLE_STYLE_MESSAGE_TYPE,
    TABLE_STYLE_PRESET_MESSAGE_TYPE,
    TABLE_STYLE_NETWORK_MESSAGE_TYPE,
    STYLESHEET_MESSAGE_TYPE,
];

/// Read merged-cell regions through the cached focused Keynote owner.
pub(super) fn regions_in_editor(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
) -> Result<Vec<Region>> {
    if editor.package().exact_source_owner().is_none() {
        return super::compatibility_slide_table_merge_regions(
            editor,
            slide_index,
            model_object_id,
        );
    }

    let (components, limits) = editor
        .package()
        .shared_component_catalog(map_catalog_error)?;
    let Some(table_position) = focused_table_position(
        editor.package(),
        components.as_ref(),
        slide_index,
        model_object_id,
    )?
    else {
        return Err(Error::ParseError(format!(
            "Keynote table model {model_object_id} is not owned by slide {slide_index}"
        )));
    };
    let options =
        litchi_keynote::ReadOptions::new(limits, litchi_keynote::SemanticLimits::default());
    let reader = MergeReader::__from_shared_catalog(components, options)
        .map_err(map_focused_merges_error)?;
    match reader.slide_table_merges(
        SlideSelector::index(slide_index),
        TableSelector::index(table_position),
    ) {
        Ok(regions) => Ok(regions),
        // Preserve the existing compatibility arm for legacy table-model
        // roles. All other post-admission failures remain typed focused
        // failures and must not widen into the archive-wide reader.
        Err(litchi_keynote::SlideTableMergesError::UnsupportedDependency) => {
            crate::numbers::editor::table_cell_merges_in_package(editor.package(), model_object_id)
        },
        Err(error) => Err(map_focused_merges_error(error)),
    }
}

/// Resolve a native model identifier to the positional selector expected by
/// the focused Keynote reader. Only rooted slide metadata and `TableInfo`
/// references are projected; a `TableModelArchive` payload is never decoded
/// by this admission scan.
fn focused_table_position(
    package: &IWorkPackage,
    components: &ComponentCatalog,
    slide_index: usize,
    model_object_id: u64,
) -> Result<Option<usize>> {
    let mut budget = SelectorBudget::new(package, components)?;
    budget.charge_components(components.iter().count())?;
    let locations = ObjectLocations::build(components, &mut budget)?;
    let root = locations.object(ROOT_OBJECT_ID)?;
    let document = unique_message(root, DOCUMENT_MESSAGE_TYPE, "KN.DocumentArchive")?;
    budget.charge_work(document.data.len().saturating_mul(8))?;
    let show_id = decode_show_identifier(package, &document.data)?;
    let show = locations.object(show_id)?;
    let show_message = unique_message(show, SHOW_MESSAGE_TYPE, "KN.ShowArchive")?;
    budget.charge_work(show_message.data.len().saturating_mul(8))?;
    let node_id = decode_slide_identifier(package, &show_message.data, slide_index)?;
    let node = locations.object(node_id)?;
    let node_message = unique_message(node, SLIDE_NODE_MESSAGE_TYPE, "KN.SlideNodeArchive")?;
    budget.charge_work(node_message.data.len().saturating_mul(8))?;
    let (slide_id, _is_skipped) = decode_slide_node(package, &node_message.data, slide_index)?;
    let slide = locations.object(slide_id)?;
    let slide_message = unique_message(slide, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
    budget.charge_work(slide_message.data.len().saturating_mul(8))?;
    let expected_drawable_references =
        precharge_drawable_references(package, &slide_message.data, &mut budget)?;
    let (owned_drawables, z_order) =
        decode_slide_drawables(package, &slide_message.data, slide_index)?;
    let drawable_references = owned_drawables
        .len()
        .checked_add(z_order.len())
        .ok_or_else(|| invalid_source("Keynote drawable reference count exceeds usize"))?;
    if drawable_references != expected_drawable_references {
        return Err(invalid_source(
            "Keynote drawable projection changed its reference count",
        ));
    }
    validate_drawable_lists(&locations, &owned_drawables, &z_order)?;

    let mut seen_models = HashSet::new();
    seen_models
        .try_reserve(z_order.len())
        .map_err(|_| allocation_error(z_order.len()))?;
    let mut table_position = 0usize;
    let mut selected = None;
    for drawable_id in z_order {
        let drawable = locations.object(drawable_id)?;
        let drawable_work = drawable
            .messages
            .iter()
            .try_fold(0usize, |total, message| {
                total.checked_add(message.data.len())
            })
            .ok_or_else(|| invalid_source("Keynote drawable payload work exceeds usize"))?;
        budget.charge_work(drawable_work.saturating_mul(8))?;
        let info_count = drawable
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .count();
        let has_role_alias = drawable.messages.iter().any(|message| {
            ROLE_MESSAGE_TYPES.contains(&message.type_) && message.type_ != TABLE_INFO_MESSAGE_TYPE
        });
        if info_count == 0 {
            if has_role_alias {
                return Err(unsupported_dependency());
            }
            continue;
        }
        if info_count != 1 || has_role_alias {
            return Err(unsupported_dependency());
        }
        let info = unique_message(drawable, TABLE_INFO_MESSAGE_TYPE, "TST.TableInfoArchive")?;
        budget.charge_references(1)?;
        let info = decode_table_info(package, &info.data, slide_id)?;
        let model_id = info.table_model().identifier().get();
        if !seen_models.insert(model_id) {
            return Err(invalid_source(
                "Keynote slide table model is rooted more than once",
            ));
        }
        // Keep model-role admission in the focused reader. In particular,
        // legacy model roles must reach that reader so its explicit
        // `UnsupportedDependency` compatibility arm remains observable.
        let _model = locations.object(model_id)?;
        if model_id == model_object_id && selected.replace(table_position).is_some() {
            return Err(invalid_source(
                "Keynote table model has multiple rooted slide owners",
            ));
        }
        table_position = table_position
            .checked_add(1)
            .ok_or_else(|| invalid_source("Keynote table position exceeds usize"))?;
    }
    Ok(selected)
}

fn validate_drawable_lists(
    locations: &ObjectLocations<'_>,
    owned_drawables: &[u64],
    z_order: &[u64],
) -> Result<()> {
    let mut owned_seen = HashSet::new();
    owned_seen
        .try_reserve(owned_drawables.len())
        .map_err(|_| allocation_error(owned_drawables.len()))?;
    for &identifier in owned_drawables {
        locations.object(identifier)?;
        if !owned_seen.insert(identifier) {
            return Err(invalid_source("Keynote slide owned drawable is repeated"));
        }
    }
    let mut z_order_seen = HashSet::new();
    z_order_seen
        .try_reserve(z_order.len())
        .map_err(|_| allocation_error(z_order.len()))?;
    for &identifier in z_order {
        locations.object(identifier)?;
        if !z_order_seen.insert(identifier) {
            return Err(invalid_source("Keynote slide drawable z-order is repeated"));
        }
        if !owned_seen.contains(&identifier) {
            return Err(invalid_source(
                "Keynote slide z-order drawable is not body-owned",
            ));
        }
    }
    Ok(())
}

fn decode_show_identifier(package: &IWorkPackage, source: &[u8]) -> Result<u64> {
    let limits = wire_limits(package, source)?;
    let recursion_limit = u32::try_from(limits.max_nesting())
        .map_err(|_| invalid_source("Keynote document nesting limit overflows u32"))?;
    keynote_document_codec::decode_show_identifier(
        source,
        keynote_document_codec::DecodeOptions::new(limits.max_input_bytes(), recursion_limit)
            .with_max_fields(limits.max_fields())
            .with_max_work_bytes(limits.max_rewrite_work()),
    )
    .map_err(map_document_error)
}

fn decode_slide_identifier(
    package: &IWorkPackage,
    source: &[u8],
    slide_index: usize,
) -> Result<u64> {
    let limits = wire_limits(package, source)?;
    let recursion_limit = u32::try_from(limits.max_nesting())
        .map_err(|_| invalid_source("Keynote show nesting limit overflows u32"))?;
    let maximum_references = source.len().clamp(1, WireLimits::MAX_INPUT_BYTES);
    keynote_show_codec::decode_slide_reference_at(
        source,
        slide_index,
        keynote_show_codec::DecodeOptions::new(
            limits.max_input_bytes(),
            maximum_references,
            recursion_limit,
        )
        .with_max_fields(limits.max_fields())
        .with_max_work_bytes(limits.max_rewrite_work()),
    )
    .map_err(map_show_error)?
    .map(|reference| reference.identifier())
    .ok_or_else(|| Error::ParseError(format!("Keynote slide index {slide_index} is out of range")))
}

fn decode_slide_node(
    package: &IWorkPackage,
    source: &[u8],
    slide_index: usize,
) -> Result<(u64, bool)> {
    let limits = wire_limits(package, source)?;
    litchi_keynote::__decode_slide_node_projection(source, limits, slide_index)
        .map_err(map_projection_error)
}

fn decode_slide_drawables(
    package: &IWorkPackage,
    source: &[u8],
    slide_index: usize,
) -> Result<(Vec<u64>, Vec<u64>)> {
    let limits = wire_limits(package, source)?;
    litchi_keynote::__decode_slide_drawable_projection(source, limits, slide_index)
        .map_err(map_projection_error)
}

/// Count repeated drawable references before the focused projection allocates
/// its owned identifier vectors. The projection validates each reference
/// again; this first pass exists solely to charge the host aggregate budget
/// before any collection growth.
fn precharge_drawable_references(
    package: &IWorkPackage,
    source: &[u8],
    budget: &mut SelectorBudget,
) -> Result<usize> {
    let limits = wire_limits(package, source)?;
    let view = WireView::parse_with_limits(source, limits)
        .map_err(|error| map_common_wire_error(error, "Keynote slide drawables"))?;
    let references = view
        .fields()
        .filter(|field| {
            matches!(
                field.number(),
                OWNED_DRAWABLE_FIELD | DRAWABLE_Z_ORDER_FIELD
            )
        })
        .count();
    budget.charge_items(references)?;
    budget.charge_references(references)?;
    Ok(references)
}

fn map_common_wire_error(error: litchi_iwa_common::Error, context: &str) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => {
            let Some(kind) = map_common_wire_limit(kind) else {
                return invalid_source(format!("{context} has an unsupported wire limit"));
            };
            Error::KeynoteSlideTableMerges(SlideTableMergesError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: limit as u64,
            })
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => allocation_error(amount),
        error => invalid_source(format!("{context} wire projection failed: {error}")),
    }
}

fn map_common_wire_limit(
    kind: litchi_iwa_common::LimitKind,
) -> Option<litchi_keynote::SlideTableMergesLimitKind> {
    Some(match kind {
        litchi_iwa_common::LimitKind::InputBytes => {
            litchi_keynote::SlideTableMergesLimitKind::InputBytes
        },
        litchi_iwa_common::LimitKind::Fields => {
            litchi_keynote::SlideTableMergesLimitKind::WireFields
        },
        litchi_iwa_common::LimitKind::OutputBytes => {
            litchi_keynote::SlideTableMergesLimitKind::OutputBytes
        },
        litchi_iwa_common::LimitKind::Nesting => {
            litchi_keynote::SlideTableMergesLimitKind::WireNesting
        },
        litchi_iwa_common::LimitKind::RewriteWork => {
            litchi_keynote::SlideTableMergesLimitKind::WireWork
        },
        litchi_iwa_common::LimitKind::TableRows
        | litchi_iwa_common::LimitKind::TableColumns
        | litchi_iwa_common::LimitKind::TableCells
        | litchi_iwa_common::LimitKind::MaterializedCells => return None,
    })
}

fn decode_table_info(
    package: &IWorkPackage,
    source: &[u8],
    slide_id: u64,
) -> Result<table_info_codec::TableInfoSnapshot> {
    let limits = wire_limits(package, source)?;
    let recursion_limit = u32::try_from(limits.max_nesting())
        .map_err(|_| invalid_source("Keynote table-info nesting limit overflows u32"))?;
    let snapshot = table_info_codec::decode_table_info_with_parent(
        source,
        table_info_codec::DecodeOptions::new(
            limits.max_input_bytes(),
            limits.max_fields(),
            limits.max_rewrite_work(),
            recursion_limit,
        ),
    )
    .map_err(map_table_info_error)?;
    if snapshot.parent() != NonZeroU64::new(slide_id) {
        return Err(invalid_source(
            "Keynote table-info parent does not match the selected slide",
        ));
    }
    Ok(snapshot)
}

fn wire_limits(package: &IWorkPackage, source: &[u8]) -> Result<WireLimits> {
    let archive_limits = package.limits().archive_limits();
    let source_bytes = source.len().max(1);
    let max_input_bytes = archive_limits
        .max_message_bytes()
        .min(archive_limits.max_archive_bytes())
        .min(package.limits().max_iwa_stream_bytes())
        .min(WireLimits::MAX_INPUT_BYTES);
    WireLimits::default()
        .with_input_bytes(source_bytes.min(max_input_bytes))
        .and_then(|limits| {
            limits.with_fields(
                source_bytes
                    .saturating_mul(2)
                    .clamp(1, WireLimits::MAX_FIELDS),
            )
        })
        .and_then(|limits| {
            limits.with_rewrite_work(
                source_bytes
                    .saturating_mul(8)
                    .clamp(1, WireLimits::MAX_REWRITE_WORK),
            )
        })
        .map_err(|error| invalid_source(format!("invalid Keynote merge selector limits: {error}")))
}

struct ObjectLocations<'catalog> {
    objects: HashMap<u64, &'catalog ArchiveObject>,
}

/// Aggregate admission ledger for the host-side identity projection.
///
/// Root and drawable projections each receive a local wire profile. Without a
/// second ledger, repeating a large slide reference or table-info payload
/// would multiply those local ceilings by the number of candidates. This
/// bounded ledger is established before the host builds any selector maps and
/// charges component, reference, payload-item, and wire-work totals once.
struct SelectorBudget {
    objects: usize,
    max_objects: usize,
    components: usize,
    max_components: usize,
    items: usize,
    max_items: usize,
    references: usize,
    max_references: usize,
    work: usize,
    max_work: usize,
}

impl SelectorBudget {
    fn new(package: &IWorkPackage, catalog: &ComponentCatalog) -> Result<Self> {
        let mut payload_bytes = 0usize;
        for component in catalog.iter() {
            for object in &component.archive().objects {
                for message in &object.messages {
                    payload_bytes = payload_bytes
                        .checked_add(message.data.len())
                        .ok_or_else(|| allocation_error(usize::MAX))?;
                }
            }
        }
        let maximum = litchi_keynote::SemanticLimits::default().max_references();
        Ok(Self {
            objects: 0,
            max_objects: WireLimits::MAX_FIELDS,
            components: 0,
            max_components: package.limits().max_entries().min(WireLimits::MAX_FIELDS),
            items: 0,
            max_items: maximum.min(WireLimits::MAX_FIELDS),
            references: 0,
            max_references: maximum.min(WireLimits::MAX_FIELDS),
            work: 0,
            max_work: payload_bytes
                .saturating_mul(16)
                .clamp(1, WireLimits::MAX_REWRITE_WORK),
        })
    }

    fn charge_objects(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.objects,
            amount,
            self.max_objects,
            litchi_keynote::SlideTableMergesLimitKind::PayloadObjects,
        )
    }

    fn charge_components(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.components,
            amount,
            self.max_components,
            litchi_keynote::SlideTableMergesLimitKind::Components,
        )
    }

    fn charge_items(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.items,
            amount,
            self.max_items,
            litchi_keynote::SlideTableMergesLimitKind::PayloadItems,
        )
    }

    fn charge_references(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.references,
            amount,
            self.max_references,
            litchi_keynote::SlideTableMergesLimitKind::References,
        )
    }

    fn charge_work(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.work,
            amount,
            self.max_work,
            litchi_keynote::SlideTableMergesLimitKind::WireWork,
        )
    }

    fn charge(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: litchi_keynote::SlideTableMergesLimitKind,
    ) -> Result<()> {
        let observed = current.saturating_add(amount);
        if observed > maximum {
            return Err(Error::KeynoteSlideTableMerges(
                SlideTableMergesError::LimitExceeded {
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
        let object_count = components.iter().try_fold(0usize, |count, component| {
            count.checked_add(component.archive().objects.len())
        });
        let object_count = object_count.ok_or_else(|| allocation_error(usize::MAX))?;
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
                        "Keynote object {identifier} occurs in multiple components"
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
            .ok_or_else(|| invalid_source(format!("Keynote object {identifier} is missing")))
    }
}

fn unique_message<'object>(
    object: &'object ArchiveObject,
    message_type: u32,
    name: &str,
) -> Result<&'object RawMessage> {
    let mut selected = None;
    for message in &object.messages {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(message).is_some() {
            return Err(invalid_source(format!("{name} has duplicate payloads")));
        }
    }
    selected.ok_or_else(|| invalid_source(format!("{name} payload is missing")))
}

fn map_focused_merges_error(error: SlideTableMergesError) -> Error {
    Error::KeynoteSlideTableMerges(error)
}

fn map_catalog_error(error: ArchiveError) -> Error {
    Error::KeynoteSlideTableMerges(match error {
        ArchiveError::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableMergesError::LimitExceeded {
            kind: map_archive_limit(kind),
            observed,
            maximum,
        },
        ArchiveError::Allocation { amount, .. } => SlideTableMergesError::Allocation { amount },
        _ => SlideTableMergesError::InvalidSource,
    })
}

fn map_archive_limit(kind: ArchiveLimitKind) -> litchi_keynote::SlideTableMergesLimitKind {
    use litchi_keynote::SlideTableMergesLimitKind;
    match kind {
        ArchiveLimitKind::InputBytes => SlideTableMergesLimitKind::InputBytes,
        ArchiveLimitKind::OutputBytes => SlideTableMergesLimitKind::OutputBytes,
        ArchiveLimitKind::Entries => SlideTableMergesLimitKind::Entries,
        ArchiveLimitKind::MemberNameBytes
        | ArchiveLimitKind::CompressedEntryBytes
        | ArchiveLimitKind::EntryBytes => SlideTableMergesLimitKind::EntryBytes,
        ArchiveLimitKind::MetadataBytes => SlideTableMergesLimitKind::TotalBytes,
        ArchiveLimitKind::TotalBytes | ArchiveLimitKind::IwaTotalBytes => {
            SlideTableMergesLimitKind::TotalBytes
        },
        ArchiveLimitKind::IwaStreamBytes => SlideTableMergesLimitKind::PayloadObjects,
    }
}

fn map_document_error(error: keynote_document_codec::DecodeError) -> Error {
    Error::KeynoteSlideTableMerges(map_wire_error(
        error.field_limit_values(),
        error.work_limit_values(),
        error.wire_resource_limit().map(|resource| match resource {
            keynote_document_codec::WireResourceLimit::Bytes { observed, maximum } => {
                WireResourceValues::Bytes { observed, maximum }
            },
            keynote_document_codec::WireResourceLimit::Nesting { observed, maximum } => {
                WireResourceValues::Nesting {
                    observed: observed as usize,
                    maximum: maximum as usize,
                }
            },
            _ => WireResourceValues::Unknown,
        }),
        None,
    ))
}

fn map_show_error(error: keynote_show_codec::DecodeError) -> Error {
    Error::KeynoteSlideTableMerges(map_wire_error(
        error.field_limit_values(),
        error.work_limit_values(),
        error.wire_resource_limit().map(|resource| match resource {
            keynote_show_codec::WireResourceLimit::Bytes { observed, maximum } => {
                WireResourceValues::Bytes { observed, maximum }
            },
            keynote_show_codec::WireResourceLimit::Nesting { observed, maximum } => {
                WireResourceValues::Nesting {
                    observed: observed as usize,
                    maximum: maximum as usize,
                }
            },
            _ => WireResourceValues::Unknown,
        }),
        error.allocation_amount(),
    ))
}

fn map_table_info_error(error: table_info_codec::DecodeError) -> Error {
    Error::KeynoteSlideTableMerges(map_wire_error(
        error.field_limit_values(),
        error.work_limit_values(),
        error.wire_resource_limit().map(|resource| match resource {
            WireResourceLimit::Bytes { observed, maximum } => WireResourceValues::Bytes {
                observed: observed.unwrap_or(usize::MAX),
                maximum: maximum.unwrap_or(usize::MAX),
            },
            WireResourceLimit::Nesting { observed, maximum } => WireResourceValues::Nesting {
                observed: observed.unwrap_or(u32::MAX) as usize,
                maximum: maximum.unwrap_or(u32::MAX) as usize,
            },
            _ => WireResourceValues::Unknown,
        }),
        error.allocation_amount(),
    ))
}

fn map_projection_error(error: litchi_keynote::ReadError) -> Error {
    Error::KeynoteSlideTableMerges(match error {
        litchi_keynote::ReadError::NotKeynote => SlideTableMergesError::UnsupportedSource,
        litchi_keynote::ReadError::Archive(ArchiveError::Limit {
            kind,
            observed,
            maximum,
        }) => SlideTableMergesError::LimitExceeded {
            kind: map_archive_limit(kind),
            observed,
            maximum,
        },
        litchi_keynote::ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableMergesError::LimitExceeded {
            kind: match kind {
                litchi_keynote::PayloadLimitKind::Bytes => {
                    litchi_keynote::SlideTableMergesLimitKind::InputBytes
                },
                litchi_keynote::PayloadLimitKind::Fields => {
                    litchi_keynote::SlideTableMergesLimitKind::WireFields
                },
                litchi_keynote::PayloadLimitKind::Nesting => {
                    litchi_keynote::SlideTableMergesLimitKind::WireNesting
                },
                litchi_keynote::PayloadLimitKind::Work => {
                    litchi_keynote::SlideTableMergesLimitKind::WireWork
                },
                _ => litchi_keynote::SlideTableMergesLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_keynote::ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableMergesError::LimitExceeded {
            kind: match kind {
                litchi_keynote::SemanticLimitKind::Objects => {
                    litchi_keynote::SlideTableMergesLimitKind::PayloadObjects
                },
                litchi_keynote::SemanticLimitKind::Slides => {
                    litchi_keynote::SlideTableMergesLimitKind::Components
                },
                litchi_keynote::SemanticLimitKind::References => {
                    litchi_keynote::SlideTableMergesLimitKind::References
                },
                litchi_keynote::SemanticLimitKind::TextStorages => {
                    litchi_keynote::SlideTableMergesLimitKind::PayloadMessages
                },
                litchi_keynote::SemanticLimitKind::TextFragments => {
                    litchi_keynote::SlideTableMergesLimitKind::PayloadItems
                },
                litchi_keynote::SemanticLimitKind::TextBytes => {
                    litchi_keynote::SlideTableMergesLimitKind::Retained
                },
                _ => litchi_keynote::SlideTableMergesLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_keynote::ReadError::Allocation { amount, .. } => {
            SlideTableMergesError::Allocation { amount }
        },
        _ => SlideTableMergesError::InvalidSource,
    })
}

fn map_wire_error(
    fields: Option<(usize, usize)>,
    work: Option<(usize, usize)>,
    resource: Option<WireResourceValues>,
    allocation: Option<usize>,
) -> SlideTableMergesError {
    if let Some(amount) = allocation {
        return SlideTableMergesError::Allocation { amount };
    }
    if let Some((observed, maximum)) = fields {
        return SlideTableMergesError::LimitExceeded {
            kind: litchi_keynote::SlideTableMergesLimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = work {
        return SlideTableMergesError::LimitExceeded {
            kind: litchi_keynote::SlideTableMergesLimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(resource) = resource {
        let (kind, observed, maximum) = match resource {
            WireResourceValues::Bytes { observed, maximum } => (
                litchi_keynote::SlideTableMergesLimitKind::InputBytes,
                observed,
                maximum,
            ),
            WireResourceValues::Nesting { observed, maximum } => (
                litchi_keynote::SlideTableMergesLimitKind::WireNesting,
                observed,
                maximum,
            ),
            WireResourceValues::Unknown => {
                return SlideTableMergesError::InvalidSource;
            },
        };
        return SlideTableMergesError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    SlideTableMergesError::InvalidSource
}

#[derive(Clone, Copy)]
enum WireResourceValues {
    Bytes { observed: usize, maximum: usize },
    Nesting { observed: usize, maximum: usize },
    Unknown,
}

fn unsupported_dependency() -> Error {
    Error::KeynoteSlideTableMerges(SlideTableMergesError::UnsupportedDependency)
}

fn allocation_error(amount: usize) -> Error {
    Error::KeynoteSlideTableMerges(SlideTableMergesError::Allocation { amount })
}

fn invalid_source(message: impl Into<String>) -> Error {
    let _ = message.into();
    Error::KeynoteSlideTableMerges(SlideTableMergesError::InvalidSource)
}
