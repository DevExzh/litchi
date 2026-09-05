//! Slide ownership and native table graph discovery.

use prost::Message;

use super::super::keynote_object_catalog::{
    KeynoteObjectCatalog, KeynoteObjectCatalogError, map_catalog_error,
};
use super::*;
use crate::protobuf::tst::TableInfoArchive;
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::table::appearance::{
    Appearance as CommonTableAppearance, Banding, GridlineVisibility, Gridlines, RowSizing,
};
use litchi_iwa_protos::{
    keynote_document_codec, keynote_show_codec, table_appearance_codec, table_info_codec,
    table_model_discovery_codec,
};

#[derive(Debug, Clone)]
pub(super) struct SlideTableGraph {
    pub(super) info: KeynoteSlideTableInfo,
    pub(super) slide_archive: String,
    pub(super) slide_component_id: u64,
}

#[derive(Debug, Clone)]
pub(super) struct CatalogSlideContext {
    pub(super) slide_id: u64,
    pub(super) slide: kn::SlideArchive,
    pub(super) focused_table_appearance_package: Option<litchi_keynote::Package>,
}

#[derive(Debug, Clone)]
struct CatalogTableModelFacts {
    name: String,
    rows: u32,
    columns: u32,
    style_identifier: u64,
    style_preset_identifier: Option<u64>,
}

pub(super) fn require_table_model(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
) -> Result<KeynoteSlideTableInfo> {
    let mut matches = editor
        .slide_tables(slide_index)?
        .into_iter()
        .filter(|table| table.model_object_id == model_object_id);
    let table = matches.next().ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote table model {model_object_id} is not owned by slide {slide_index}"
        ))
    })?;
    if matches.next().is_some() {
        return Err(Error::ParseError(format!(
            "Keynote table model {model_object_id} is owned by slide {slide_index} more than once"
        )));
    }
    Ok(table)
}

/// Resolve one table through the bounded catalog used by slide listing.
///
/// Keeping one admission path for listing and mutation prevents a private
/// caller from observing a different model candidate after the public listing
/// has already established ownership.  In particular, a valid legacy 6000
/// model remains supported when it is the sole candidate, while malformed or
/// mixed current/legacy candidates are rejected by the catalog's strict role
/// census rather than being retried through an eager graph.
pub(super) fn slide_table_graph(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<SlideTableGraph> {
    let mut catalog = KeynoteObjectCatalog::build(editor.package()).map_err(map_catalog_error)?;
    let context = catalog_slide_context(editor.package(), &mut catalog, slide_index)?;
    slide_table_graph_from_catalog_context(
        editor,
        &mut catalog,
        slide_index,
        drawable_object_id,
        &context,
    )
}

/// Return the table's zero-based position in a catalog-backed slide listing.
///
/// The catalog stores only message descriptors, so this remains a bounded
/// metadata lookup and does not decode another table payload merely to map a
/// legacy drawable identity to the focused selector.
fn catalog_table_position(
    catalog: &KeynoteObjectCatalog,
    references: &[tsp::Reference],
    drawable_object_id: u64,
) -> Result<usize> {
    let mut table_position = None;
    let mut table_count = 0usize;
    for reference in references {
        let is_table = catalog
            .message_type_count(reference.identifier, TABLE_INFO_MESSAGE_TYPE)
            .map_err(map_catalog_error)?
            > 0;
        if !is_table {
            continue;
        }
        if reference.identifier == drawable_object_id
            && table_position.replace(table_count).is_some()
        {
            return Err(Error::ParseError(format!(
                "Keynote slide table {drawable_object_id} has ambiguous z-order position"
            )));
        }
        table_count = table_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("Keynote slide table count exceeds usize".to_owned())
        })?;
    }
    table_position.ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide table {drawable_object_id} has no z-order position"
        ))
    })
}

/// Build the focused package view once for a catalog-backed slide listing.
///
/// The focused package is immutable and internally shares its physical source
/// through an `Arc`, so retaining this per-slide view avoids reparsing the ZIP
/// and rebuilding its object index for every table in the same listing.
fn focused_table_appearance_package(
    package: &IWorkPackage,
) -> Result<Option<litchi_keynote::Package>> {
    if !package.source_is_exact() {
        return Ok(None);
    }
    let source = package.exact_source_bytes().ok_or_else(|| {
        Error::InvalidFormat("focused Keynote table appearance source is not exact".to_owned())
    })?;
    let focused = litchi_keynote::Package::from_bytes(source).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote table appearance source failed: {error}"
        ))
    })?;
    if focused
        .__is_litchi_source_built_compatibility()
        .map_err(|error| {
            Error::InvalidFormat(format!(
                "focused Keynote compatibility classification failed: {error}"
            ))
        })?
    {
        return Ok(None);
    }
    Ok(Some(focused))
}

/// Read appearance through the focused Keynote package owner for an admitted
/// exact physical source. The source provenance check above keeps synthetic
/// and compatibility snapshots on their existing catalog path.
fn focused_table_appearance(
    package: &litchi_keynote::Package,
    slide_index: usize,
    table_position: usize,
) -> Result<CommonTableAppearance> {
    package
        .slide_table_appearance(
            litchi_keynote::SlideSelector::index(slide_index),
            litchi_keynote::TableSelector::index(table_position),
        )
        .map_err(|error| {
            Error::InvalidFormat(format!(
                "focused Keynote table appearance read failed: {error}"
            ))
        })
}

/// Resolve the root slide objects through the bounded catalog.
///
/// The document's show edge and the selected show slide edge use the focused
/// generated-free projections. The returned slide value is deliberately
/// short-lived and owns only the selected slide projection. The catalog
/// retains object slots and message descriptors, not the package's decompressed
/// payloads.
pub(super) fn catalog_slide_context(
    package: &IWorkPackage,
    catalog: &mut KeynoteObjectCatalog,
    slide_index: usize,
) -> Result<CatalogSlideContext> {
    let show_identifier = catalog
        .with_message_data_type(
            package,
            1,
            DOCUMENT_MESSAGE_TYPE,
            "KN.DocumentArchive",
            decode_catalog_document_show_identifier,
        )
        .map_err(map_catalog_error)?;
    let node_identifier = catalog
        .with_message_data_type(
            package,
            show_identifier,
            SHOW_MESSAGE_TYPE,
            "KN.ShowArchive",
            |source| decode_catalog_slide_identifier(source, slide_index),
        )
        .map_err(map_catalog_error)?;
    let slide_id = catalog
        .with_message_data_type(
            package,
            node_identifier,
            4,
            "KN.SlideNodeArchive",
            |source| decode_catalog_slide_node_identifier(source, slide_index, package),
        )
        .map_err(map_catalog_error)?;
    let slide: kn::SlideArchive = catalog
        .decode_type(package, slide_id, 5, "KN.SlideArchive")
        .map_err(map_catalog_error)?;
    let focused_package = focused_table_appearance_package(package)?;
    Ok(CatalogSlideContext {
        slide_id,
        slide,
        focused_table_appearance_package: focused_package,
    })
}

fn decode_catalog_slide_node_identifier(
    source: &[u8],
    slide_index: usize,
    package: &IWorkPackage,
) -> std::result::Result<u64, KeynoteObjectCatalogError> {
    let wire_limits = keynote_slide_node_wire_limits(package.limits(), source)?;
    litchi_keynote::__decode_slide_node_projection(source, wire_limits, slide_index)
        .map(|(identifier, _is_skipped)| identifier)
        .map_err(|error| {
            KeynoteObjectCatalogError::InvalidSource(format!(
                "malformed KN.SlideNodeArchive payload: {error}"
            ))
        })
}

fn decode_catalog_document_show_identifier(
    source: &[u8],
) -> std::result::Result<u64, KeynoteObjectCatalogError> {
    // This host projection only needs the required KN/TSA envelope and show
    // edge. Optional document graph tables stay opaque to preserve the
    // bounded lazy path's ownership reduction.
    let options = keynote_document_options(source);
    keynote_document_codec::decode_template_identifier(source, options).map_err(|error| {
        KeynoteObjectCatalogError::InvalidSource(format!(
            "malformed KN.DocumentArchive payload: {error}"
        ))
    })?;
    keynote_document_codec::decode_show_identifier(source, options).map_err(|error| {
        KeynoteObjectCatalogError::InvalidSource(format!(
            "malformed KN.DocumentArchive payload: {error}"
        ))
    })
}

fn decode_catalog_slide_identifier(
    source: &[u8],
    slide_index: usize,
) -> std::result::Result<u64, KeynoteObjectCatalogError> {
    let reference = keynote_show_codec::decode_slide_reference_at(
        source,
        slide_index,
        keynote_show_options(source),
    )
    .map_err(|error| {
        KeynoteObjectCatalogError::InvalidSource(format!(
            "malformed KN.ShowArchive payload: {error}"
        ))
    })?;
    reference
        .map(|reference| reference.identifier())
        .ok_or_else(|| {
            KeynoteObjectCatalogError::InvalidSource(format!(
                "Keynote slide index {slide_index} is out of range in KN.ShowArchive"
            ))
        })
}

/// Resolve one table using an already-decoded catalog slide context.
///
/// Listing passes this context for every drawable on the selected slide.  It
/// is intentionally a separate helper so the root/show lazy projections and
/// node/slide semantic decodes remain constant when a slide contains many
/// tables.
pub(super) fn slide_table_graph_from_catalog_context(
    editor: &KeynoteEditor,
    catalog: &mut KeynoteObjectCatalog,
    slide_index: usize,
    drawable_object_id: u64,
    context: &CatalogSlideContext,
) -> Result<SlideTableGraph> {
    let package = editor.package();
    for (name, references) in [
        ("owned_drawables", &context.slide.owned_drawables),
        ("drawables_z_order", &context.slide.drawables_z_order),
    ] {
        if references
            .iter()
            .filter(|reference| reference.identifier == drawable_object_id)
            .count()
            != 1
        {
            return Err(Error::ParseError(format!(
                "Keynote slide {} {name} does not own table {drawable_object_id} exactly once",
                context.slide_id
            )));
        }
    }
    validate_catalog_table_info_role(catalog, drawable_object_id)?;
    let (table_info, table_info_projection) = catalog
        .with_message_data_type(
            package,
            drawable_object_id,
            TABLE_INFO_MESSAGE_TYPE,
            "TableInfoArchive",
            decode_catalog_table_info,
        )
        .map_err(map_catalog_error)?;
    if table_info
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(context.slide_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote table {drawable_object_id} does not name slide {} as its parent",
            context.slide_id
        )));
    }
    let model_id = table_info_projection.table_model().identifier().get();
    let model = catalog_table_model_facts(package, catalog, model_id)?;
    let table_position = catalog_table_position(
        catalog,
        &context.slide.drawables_z_order,
        drawable_object_id,
    )?;
    let slide_archive = catalog
        .archive_name(context.slide_id)
        .map_err(map_catalog_error)?
        .to_owned();
    let slide_component_id =
        component_identifier_for_entry(package, &slide_archive)?.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide component {slide_archive} is not registered"
            ))
        })?;
    let lock_state = TableLockState::from_locked(table_info_projection.locked().unwrap_or(false));
    let appearance = if let Some(focused) = context.focused_table_appearance_package.as_ref() {
        focused_table_appearance(focused, slide_index, table_position)?
    } else {
        catalog_table_appearance(
            package,
            catalog,
            model.style_identifier,
            model.style_preset_identifier,
        )?
    };
    Ok(SlideTableGraph {
        info: KeynoteSlideTableInfo {
            slide_index,
            slide_id: context.slide_id,
            drawable_object_id,
            model_object_id: model_id,
            name: model.name,
            rows: model.rows as usize,
            columns: model.columns as usize,
            geometry: crate::shapes::geometry_from_drawable(&table_info.super_)?,
            appearance,
            lock_state,
        },
        slide_archive,
        slide_component_id,
    })
}

fn decode_catalog_table_info(
    source: &[u8],
) -> std::result::Result<
    (TableInfoArchive, table_info_codec::TableInfoSnapshot),
    KeynoteObjectCatalogError,
> {
    let projected = table_info_codec::decode_table_info(source, table_info_options(source))
        .map_err(|error| KeynoteObjectCatalogError::InvalidSource(error.to_string()))?;
    let generated = TableInfoArchive::decode(source).map_err(|error| {
        KeynoteObjectCatalogError::InvalidSource(format!(
            "malformed TableInfoArchive payload: {error}"
        ))
    })?;
    Ok((generated, projected))
}

fn validate_catalog_table_info_role(catalog: &KeynoteObjectCatalog, info_id: u64) -> Result<()> {
    if catalog
        .message_type_count(info_id, TABLE_INFO_MESSAGE_TYPE)
        .map_err(map_catalog_error)?
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote table-info object {info_id} must contain exactly one table-info payload"
        )));
    }
    if catalog
        .message_type_count(info_id, 6_001)
        .map_err(map_catalog_error)?
        != 0
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote table-info object {info_id} contains a table-model role alias"
        )));
    }
    if catalog
        .message_type_count(info_id, TABLE_MODEL_ROLE_ALIAS_MESSAGE_TYPE)
        .map_err(map_catalog_error)?
        != 0
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote table-info object {info_id} contains a historical table-model role alias"
        )));
    }
    Ok(())
}

fn catalog_table_model_facts(
    package: &IWorkPackage,
    catalog: &mut KeynoteObjectCatalog,
    model_id: u64,
) -> Result<CatalogTableModelFacts> {
    let canonical_count = catalog
        .message_type_count(model_id, 6_001)
        .map_err(map_catalog_error)?;
    let legacy_count = catalog
        .message_type_count(model_id, 6_000)
        .map_err(map_catalog_error)?;
    for role in [
        TABLE_MODEL_ROLE_ALIAS_MESSAGE_TYPE,
        TABLE_STYLE_MESSAGE_TYPE,
        TABLE_STYLE_PRESET_MESSAGE_TYPE,
        TABLE_STYLE_NETWORK_MESSAGE_TYPE,
    ] {
        if catalog
            .message_type_count(model_id, role)
            .map_err(map_catalog_error)?
            != 0
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote table model {model_id} contains a model or appearance role alias"
            )));
        }
    }
    let message_type = match (canonical_count, legacy_count) {
        (0, 1) => 6_000,
        (1, 0) => 6_001,
        (0, 0) => {
            return Err(Error::InvalidFormat(format!(
                "Keynote table model {model_id} has no canonical table-model payload"
            )));
        },
        (1, 1) => {
            return Err(Error::InvalidFormat(format!(
                "Keynote table model {model_id} has both canonical and legacy payloads"
            )));
        },
        (canonical, legacy) if canonical > 1 || legacy > 1 => {
            return Err(Error::InvalidFormat(format!(
                "Keynote table model {model_id} repeats a table-model payload"
            )));
        },
        _ => unreachable!("table model message counts are nonnegative"),
    };
    catalog
        .with_message_data_type(
            package,
            model_id,
            message_type,
            "TableModelArchive",
            |source| {
                let facts = table_model_discovery_codec::decode_table_model(
                    source,
                    table_model_options(source),
                )
                .map_err(|error| KeynoteObjectCatalogError::InvalidSource(error.to_string()))?;
                let appearance_model = table_appearance_codec::decode_table_model(
                    source,
                    table_appearance_options(source),
                )
                .map_err(|error| KeynoteObjectCatalogError::InvalidSource(error.to_string()))?;
                let table_name = facts.table_name();
                let mut name = String::new();
                name.try_reserve_exact(table_name.len()).map_err(|_| {
                    KeynoteObjectCatalogError::Allocation {
                        resource: "Keynote table name",
                        amount: table_name.len(),
                    }
                })?;
                name.push_str(table_name);
                Ok(CatalogTableModelFacts {
                    name,
                    rows: facts.number_of_rows(),
                    columns: facts.number_of_columns(),
                    style_identifier: appearance_model.style_identifier(),
                    style_preset_identifier: appearance_model.style_preset_identifier(),
                })
            },
        )
        .map_err(map_catalog_error)
}

const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const MAX_CATALOG_STYLE_INHERITANCE_DEPTH: usize = 64;

/// Resolve a table's appearance using the already-built object catalog.
///
/// The older `crate::table_appearance::table_appearance` helper intentionally
/// returns owned archives for its compatibility callers.  Listing is a much
/// hotter path: calling it for every drawable would clone the cached archive
/// for the model, then scan every IWA member again for each style hop.  Keep
/// this path catalog-backed and borrow each selected payload only for the
/// duration of its strict codec callback.  The catalog therefore remains the
/// single package scan and no parsed archive is retained by this operation.
fn catalog_table_appearance(
    package: &IWorkPackage,
    catalog: &mut KeynoteObjectCatalog,
    style_identifier: u64,
    style_preset_identifier: Option<u64>,
) -> Result<CommonTableAppearance> {
    let Some(first_style_identifier) = catalog_effective_style_identifier(
        package,
        catalog,
        style_identifier,
        style_preset_identifier,
    )?
    else {
        return Ok(CommonTableAppearance::default());
    };

    let mut visited = [0_u64; MAX_CATALOG_STYLE_INHERITANCE_DEPTH];
    let mut current = Some(first_style_identifier);
    let mut row_banding = None;
    let mut row_sizing = None;
    let mut body_horizontal = None;
    let mut body_vertical = None;
    let mut header_columns_horizontal = None;
    let mut header_rows_vertical = None;
    let mut footer_rows_vertical = None;

    for visited_len in 0..=MAX_CATALOG_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = current else {
            return Ok(catalog_appearance_from_overrides(
                row_banding,
                row_sizing,
                body_horizontal,
                body_vertical,
                header_columns_horizontal,
                header_rows_vertical,
                footer_rows_vertical,
            ));
        };
        if visited[..visited_len].contains(&identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork table style inheritance cycles at {identifier}"
            )));
        }
        if visited_len == visited.len() {
            return Err(Error::InvalidFormat(format!(
                "iWork table style inheritance exceeds {MAX_CATALOG_STYLE_INHERITANCE_DEPTH} levels"
            )));
        }
        visited[visited_len] = identifier;

        validate_catalog_appearance_role(
            catalog,
            identifier,
            TABLE_STYLE_MESSAGE_TYPE,
            "table style",
        )?;

        let (parent_identifier, overrides) = catalog
            .with_message_data_type(
                package,
                identifier,
                TABLE_STYLE_MESSAGE_TYPE,
                "TableStyleArchive",
                |source| {
                    let style = table_appearance_codec::decode_table_style(
                        source,
                        table_appearance_options(source),
                    )
                    .map_err(|error| KeynoteObjectCatalogError::InvalidSource(error.to_string()))?;
                    Ok((style.parent_identifier(), style.overrides()))
                },
            )
            .map_err(map_catalog_error)?;

        row_banding = row_banding.or(overrides.row_banding);
        row_sizing = row_sizing.or(overrides.row_sizing);
        body_horizontal = body_horizontal.or(overrides.body_horizontal);
        body_vertical = body_vertical.or(overrides.body_vertical);
        header_columns_horizontal =
            header_columns_horizontal.or(overrides.header_columns_horizontal);
        header_rows_vertical = header_rows_vertical.or(overrides.header_rows_vertical);
        footer_rows_vertical = footer_rows_vertical.or(overrides.footer_rows_vertical);

        current = parent_identifier.filter(|identifier| *identifier != 0);
    }

    Err(Error::InvalidFormat(format!(
        "iWork table style inheritance exceeds {MAX_CATALOG_STYLE_INHERITANCE_DEPTH} levels"
    )))
}

fn catalog_effective_style_identifier(
    package: &IWorkPackage,
    catalog: &mut KeynoteObjectCatalog,
    style_identifier: u64,
    style_preset_identifier: Option<u64>,
) -> Result<Option<u64>> {
    if style_identifier != 0 {
        // Match the legacy resolver: a concrete model style wins over the
        // optional preset, so an unused malformed preset cannot poison it.
        return Ok(Some(style_identifier));
    }
    let Some(preset_identifier) = style_preset_identifier else {
        return Ok(None);
    };
    validate_catalog_appearance_role(
        catalog,
        preset_identifier,
        TABLE_STYLE_PRESET_MESSAGE_TYPE,
        "table style preset",
    )?;
    let network_identifier = catalog
        .with_message_data_type(
            package,
            preset_identifier,
            TABLE_STYLE_PRESET_MESSAGE_TYPE,
            "TableStylePresetArchive",
            |source| {
                let preset = table_appearance_codec::decode_table_style_preset(
                    source,
                    table_appearance_options(source),
                )
                .map_err(|error| KeynoteObjectCatalogError::InvalidSource(error.to_string()))?;
                Ok(preset.style_network_identifier())
            },
        )
        .map_err(map_catalog_error)?
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table style preset {preset_identifier} has no style network"
            ))
        })?;
    validate_catalog_appearance_role(
        catalog,
        network_identifier,
        TABLE_STYLE_NETWORK_MESSAGE_TYPE,
        "table style network",
    )?;
    let table_style_identifier = catalog
        .with_message_data_type(
            package,
            network_identifier,
            TABLE_STYLE_NETWORK_MESSAGE_TYPE,
            "TableStyleNetworkArchive",
            |source| {
                let network = table_appearance_codec::decode_table_style_network(
                    source,
                    table_appearance_options(source),
                )
                .map_err(|error| KeynoteObjectCatalogError::InvalidSource(error.to_string()))?;
                Ok(network.table_style_identifier())
            },
        )
        .map_err(map_catalog_error)?;
    if table_style_identifier == 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork table style network {network_identifier} has no table style"
        )));
    }
    Ok(Some(table_style_identifier))
}

fn validate_catalog_appearance_role(
    catalog: &KeynoteObjectCatalog,
    identifier: u64,
    expected_type: u32,
    expected_name: &str,
) -> Result<()> {
    if catalog
        .message_type_count(identifier, expected_type)
        .map_err(map_catalog_error)?
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "iWork {expected_name} {identifier} must contain exactly one role payload"
        )));
    }
    for role in [
        TABLE_STYLE_MESSAGE_TYPE,
        TABLE_STYLE_PRESET_MESSAGE_TYPE,
        TABLE_STYLE_NETWORK_MESSAGE_TYPE,
        TABLE_MODEL_MESSAGE_TYPE,
    ] {
        if role != expected_type
            && catalog
                .message_type_count(identifier, role)
                .map_err(map_catalog_error)?
                != 0
        {
            return Err(Error::InvalidFormat(format!(
                "iWork {expected_name} {identifier} contains an appearance role alias"
            )));
        }
    }
    Ok(())
}

fn catalog_appearance_from_overrides(
    row_banding: Option<bool>,
    row_sizing: Option<bool>,
    body_horizontal: Option<bool>,
    body_vertical: Option<bool>,
    header_columns_horizontal: Option<bool>,
    header_rows_vertical: Option<bool>,
    footer_rows_vertical: Option<bool>,
) -> CommonTableAppearance {
    CommonTableAppearance {
        row_banding: if row_banding.unwrap_or(false) {
            Banding::Enabled
        } else {
            Banding::Disabled
        },
        row_sizing: if row_sizing.unwrap_or(false) {
            RowSizing::FitCellContents
        } else {
            RowSizing::Fixed
        },
        gridlines: Gridlines {
            body_horizontal: if body_horizontal.unwrap_or(true) {
                GridlineVisibility::Visible
            } else {
                GridlineVisibility::Hidden
            },
            header_columns_horizontal: if header_columns_horizontal.unwrap_or(true) {
                GridlineVisibility::Visible
            } else {
                GridlineVisibility::Hidden
            },
            body_vertical: if body_vertical.unwrap_or(true) {
                GridlineVisibility::Visible
            } else {
                GridlineVisibility::Hidden
            },
            header_rows_vertical: if header_rows_vertical.unwrap_or(true) {
                GridlineVisibility::Visible
            } else {
                GridlineVisibility::Hidden
            },
            footer_rows_vertical: if footer_rows_vertical.unwrap_or(true) {
                GridlineVisibility::Visible
            } else {
                GridlineVisibility::Hidden
            },
        },
    }
}

fn table_appearance_options(source: &[u8]) -> table_appearance_codec::DecodeOptions {
    table_appearance_codec::DecodeOptions::for_source(source)
}

fn table_info_options(source: &[u8]) -> table_info_codec::DecodeOptions {
    table_info_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(4)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        64,
    )
}

fn keynote_document_options(source: &[u8]) -> keynote_document_codec::DecodeOptions {
    let source_bytes = source.len().clamp(1, WireLimits::MAX_INPUT_BYTES);
    keynote_document_codec::DecodeOptions::new(source_bytes, 64)
        .with_max_fields(source_bytes.min(WireLimits::MAX_FIELDS))
        .with_max_work_bytes(
            source_bytes
                .saturating_mul(8)
                .clamp(1, WireLimits::MAX_REWRITE_WORK),
        )
}

fn keynote_show_options(source: &[u8]) -> keynote_show_codec::DecodeOptions {
    let source_bytes = source.len().clamp(1, WireLimits::MAX_INPUT_BYTES);
    keynote_show_codec::DecodeOptions::new(source_bytes, source_bytes, 64)
        .with_max_fields(source_bytes.min(WireLimits::MAX_FIELDS))
        .with_max_work_bytes(
            source_bytes
                .saturating_mul(8)
                .clamp(1, WireLimits::MAX_REWRITE_WORK),
        )
}

fn keynote_slide_node_wire_limits(
    limits: crate::package::PackageLimits,
    source: &[u8],
) -> std::result::Result<WireLimits, KeynoteObjectCatalogError> {
    let archive_limits = limits.archive_limits();
    let source_bytes = source.len().max(1);
    let max_input_bytes = archive_limits
        .max_message_bytes()
        .min(archive_limits.max_archive_bytes())
        .min(limits.max_iwa_stream_bytes())
        .min(WireLimits::MAX_INPUT_BYTES);
    let input_bytes = source_bytes.min(max_input_bytes);
    // Archive header budgets govern framing metadata, not protobuf payloads.
    // Keep a separate source-sized field profile and the common bounded depth.
    let fields = source_bytes.min(WireLimits::MAX_FIELDS);
    let rewrite_work = source_bytes
        .saturating_mul(8)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    WireLimits::default()
        .with_input_bytes(input_bytes)
        .and_then(|limits| limits.with_fields(fields))
        .and_then(|limits| limits.with_rewrite_work(rewrite_work))
        .map_err(|error| {
            KeynoteObjectCatalogError::InvalidSource(format!(
                "invalid Keynote slide-node wire limits: {error}"
            ))
        })
}

fn table_model_options(source: &[u8]) -> table_model_discovery_codec::DecodeOptions {
    table_model_discovery_codec::DecodeOptions::for_source(source)
}

/// Find a table template through the bounded catalog.
pub(super) fn table_template_from_catalog(
    package: &IWorkPackage,
    catalog: &mut KeynoteObjectCatalog,
) -> Result<(u64, u64)> {
    let mut candidate = None;
    for object_index in 0..catalog.object_count() {
        let info_id = catalog
            .object_identifier_at(object_index)
            .map_err(map_catalog_error)?;
        if catalog
            .message_type_count(info_id, TABLE_INFO_MESSAGE_TYPE)
            .map_err(map_catalog_error)?
            == 0
        {
            continue;
        }
        validate_catalog_table_info_role(catalog, info_id)?;
        let (info, projection) = catalog
            .with_message_data_type(
                package,
                info_id,
                TABLE_INFO_MESSAGE_TYPE,
                "TableInfoArchive",
                decode_catalog_table_info,
            )
            .map_err(map_catalog_error)?;
        let model_id = projection.table_model().identifier().get();
        if info.table_model.identifier != model_id {
            return Err(Error::InvalidFormat(
                "Keynote table-info model reference projection disagrees with generated payload"
                    .to_owned(),
            ));
        }
        // The creation scaffold is deliberately detached from presentation
        // content.  Existing slide-owned tables are valid listing entries,
        // but are not additional creation templates.  Decode and validate
        // every TableInfo candidate above before applying this ownership
        // filter so malformed role payloads still fail closed.
        if info.super_.parent.is_some() {
            continue;
        }
        catalog_table_model_facts(package, catalog, model_id)?;
        if candidate.replace((info_id, model_id)).is_some() {
            return Err(Error::InvalidFormat(
                "Keynote package has multiple native table creation templates".to_owned(),
            ));
        }
    }
    candidate.ok_or_else(|| {
        Error::InvalidFormat("Keynote package has no native table creation template".to_owned())
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject, RawMessage};
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::protobuf::{tsp, tss, tst};
    use prost::Message;

    const MODEL_TYPE: u32 = 6_001;
    const LEGACY_MODEL_TYPE: u32 = 6_000;

    fn raw_object(identifier: u64, messages: Vec<(u32, Vec<u8>)>) -> ArchiveObject {
        ArchiveObject::new(
            identifier,
            messages
                .into_iter()
                .map(|(type_, data)| RawMessage { type_, data })
                .collect(),
        )
        .expect("synthetic object")
    }

    fn model_payload() -> Vec<u8> {
        let mut output = Vec::new();
        bytes_field(1, b"id", &mut output);
        bytes_field(3, &reference_payload(1), &mut output);
        bytes_field(4, &[0x0a, 0x00], &mut output);
        varint_field(6, 3, &mut output);
        varint_field(7, 4, &mut output);
        bytes_field(8, b"Table", &mut output);
        fixed64_field(16, &mut output);
        fixed64_field(17, &mut output);
        for field in 18..=21 {
            bytes_field(
                field,
                &reference_payload(u64::from(field - 17)),
                &mut output,
            );
        }
        for field in 24..=27 {
            bytes_field(
                field,
                &reference_payload(u64::from(field - 23)),
                &mut output,
            );
        }
        output
    }

    fn table_info_payload(model_identifier: u64) -> Vec<u8> {
        let mut output = vec![0x0a, 0x00, 0x12, 0x01];
        append_varint(model_identifier, &mut output);
        output
    }

    fn model_package(messages: Vec<(u32, Vec<u8>)>) -> IWorkPackage {
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/CalculationEngine.iwa",
                &Archive {
                    objects: vec![raw_object(42, messages)],
                },
            )
            .expect("synthetic archive");
        package
    }

    fn table_template_package(model_messages: Vec<(u32, Vec<u8>)>) -> IWorkPackage {
        table_template_package_with_info_messages(
            vec![(TABLE_INFO_MESSAGE_TYPE, table_info_payload(42))],
            model_messages,
        )
    }

    fn table_template_package_with_info_messages(
        info_messages: Vec<(u32, Vec<u8>)>,
        model_messages: Vec<(u32, Vec<u8>)>,
    ) -> IWorkPackage {
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/CalculationEngine.iwa",
                &Archive {
                    objects: vec![
                        raw_object(10, info_messages),
                        raw_object(42, model_messages),
                    ],
                },
            )
            .expect("synthetic archive");
        package
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn reference_payload(identifier: u64) -> Vec<u8> {
        let mut output = Vec::new();
        varint_field(1, identifier, &mut output);
        output
    }

    fn slide_node_payload(identifier: u64, is_skipped: bool) -> Vec<u8> {
        let mut output = Vec::new();
        bytes_field(2, &reference_payload(identifier), &mut output);
        varint_field(4, u64::from(is_skipped), &mut output);
        varint_field(6, 0, &mut output);
        varint_field(7, 0, &mut output);
        output
    }

    #[test]
    fn catalog_slide_node_projection_matches_generated_identifier_without_retaining_source() {
        let package = IWorkPackage::new();
        let mut source = slide_node_payload(88, true);
        let generated = kn::SlideNodeArchive::decode(source.as_slice()).expect("generated node");
        let projected = decode_catalog_slide_node_identifier(&source, 3, &package)
            .expect("focused slide-node projection");
        assert_eq!(
            projected,
            generated.slide.expect("slide reference").identifier
        );

        // The catalog keeps the projected scalar, so its result remains valid
        // after the source buffer is reused by the caller.
        source.fill(0);
        assert_eq!(projected, 88);

        let alternate = slide_node_payload(99, false);
        assert_eq!(
            decode_catalog_slide_node_identifier(&alternate, 4, &package)
                .expect("alternate focused projection"),
            99
        );
    }

    #[test]
    fn catalog_slide_node_projection_rejects_missing_duplicate_and_malformed_references() {
        let package = IWorkPackage::new();
        let reference = reference_payload(88);
        let mut missing_required = Vec::new();
        bytes_field(2, &reference, &mut missing_required);
        assert!(decode_catalog_slide_node_identifier(&missing_required, 0, &package).is_err());

        let mut duplicate = slide_node_payload(88, false);
        bytes_field(2, &reference_payload(89), &mut duplicate);
        assert!(decode_catalog_slide_node_identifier(&duplicate, 0, &package).is_err());

        let mut malformed_reference = Vec::new();
        bytes_field(2, &[0x08, 0x80, 0x00], &mut malformed_reference);
        varint_field(4, 0, &mut malformed_reference);
        varint_field(6, 0, &mut malformed_reference);
        varint_field(7, 0, &mut malformed_reference);
        assert!(decode_catalog_slide_node_identifier(&malformed_reference, 0, &package).is_err());
    }

    #[test]
    fn slide_node_payload_limits_are_independent_of_archive_header_budgets() {
        let source = slide_node_payload(88, false);
        let archive_limits = crate::package::PackageLimits::default()
            .archive_limits()
            .with_header_fields(1)
            .unwrap()
            .with_header_nesting(1)
            .unwrap();
        let limits = crate::package::PackageLimits::default()
            .with_archive_limits(archive_limits)
            .unwrap();
        let wire = keynote_slide_node_wire_limits(limits, &source).unwrap();
        assert_eq!(
            litchi_keynote::__decode_slide_node_projection(&source, wire, 0)
                .unwrap()
                .0,
            88
        );
        let limits = limits
            .with_archive_limits(archive_limits.with_message_bytes(source.len() - 1).unwrap())
            .unwrap();
        let wire = keynote_slide_node_wire_limits(limits, &source).unwrap();
        assert!(litchi_keynote::__decode_slide_node_projection(&source, wire, 0).is_err());
    }

    #[test]
    fn catalog_root_edges_use_bounded_generated_free_projections() {
        let show_reference = reference_payload(77);
        let mut document = Vec::new();
        // The document projection requires the KN/TSA super envelopes before
        // reading the show edge.  Keep this synthetic payload at that minimal
        // valid boundary so the test exercises the same contract as a real
        // document while leaving optional graph tables opaque.
        bytes_field(3, &[0x0a, 0x00], &mut document);
        bytes_field(2, &show_reference, &mut document);
        assert_eq!(
            decode_catalog_document_show_identifier(&document).expect("show identifier"),
            77
        );

        let mut missing_base = Vec::new();
        bytes_field(2, &show_reference, &mut missing_base);
        assert!(decode_catalog_document_show_identifier(&missing_base).is_err());

        let mut malformed_base = Vec::new();
        bytes_field(3, &[], &mut malformed_base);
        bytes_field(2, &show_reference, &mut malformed_base);
        assert!(decode_catalog_document_show_identifier(&malformed_base).is_err());

        let show = kn::ShowArchive {
            theme: reference(90),
            slide_tree: kn::SlideTreeArchive {
                slides: vec![reference(88)],
                ..Default::default()
            },
            size: tsp::Size {
                width: 1_024.0,
                height: 768.0,
            },
            stylesheet: reference(91),
            ..Default::default()
        }
        .encode_to_vec();
        assert_eq!(
            decode_catalog_slide_identifier(&show, 0).expect("selected slide identifier"),
            88
        );
        assert!(
            decode_catalog_slide_identifier(&show, 1)
                .expect_err("out-of-range slide")
                .to_string()
                .contains("out of range")
        );

        let mut duplicate_document = document.clone();
        bytes_field(2, &show_reference, &mut duplicate_document);
        assert!(decode_catalog_document_show_identifier(&duplicate_document).is_err());
    }

    fn appearance_model_payload(
        style_identifier: u64,
        style_preset_identifier: Option<u64>,
    ) -> Vec<u8> {
        let mut output = Vec::new();
        bytes_field(1, b"model", &mut output);
        bytes_field(3, &reference_payload(style_identifier), &mut output);
        if let Some(identifier) = style_preset_identifier {
            bytes_field(48, &reference_payload(identifier), &mut output);
        }
        // DataStore is a required nested envelope.  Discovery only needs its
        // presence, so the smallest valid empty envelope is sufficient here.
        bytes_field(4, &[0x0a, 0x00], &mut output);
        varint_field(6, 3, &mut output);
        varint_field(7, 4, &mut output);
        bytes_field(8, b"Table", &mut output);
        fixed64_field(16, &mut output);
        fixed64_field(17, &mut output);
        for field in 18..=21 {
            bytes_field(
                field,
                &reference_payload(u64::from(field - 17)),
                &mut output,
            );
        }
        for field in 24..=27 {
            bytes_field(
                field,
                &reference_payload(u64::from(field - 23)),
                &mut output,
            );
        }
        output
    }

    fn appearance_model_object(
        identifier: u64,
        style_identifier: u64,
        style_preset_identifier: Option<u64>,
    ) -> ArchiveObject {
        raw_object(
            identifier,
            vec![(
                MODEL_TYPE,
                appearance_model_payload(style_identifier, style_preset_identifier),
            )],
        )
    }

    fn appearance_style_payload(
        identifier: u64,
        parent_identifier: Option<u64>,
        overrides: [Option<bool>; 7],
        with_unknown_group: bool,
    ) -> Vec<u8> {
        let properties =
            overrides
                .iter()
                .any(Option::is_some)
                .then(|| tst::TableStylePropertiesArchive {
                    banded_rows: overrides[0],
                    auto_resize: overrides[1],
                    v_strokes_visible: overrides[2],
                    h_strokes_visible: overrides[3],
                    table_hc_divider_visible: overrides[4],
                    table_hr_divider_visible: overrides[5],
                    table_footer_divider_visible: overrides[6],
                    ..Default::default()
                });
        let mut output = tst::TableStyleArchive {
            super_: tss::StyleArchive {
                style_identifier: Some(format!("style-{identifier}")),
                parent: parent_identifier.map(reference),
                ..Default::default()
            },
            table_properties: properties,
            ..Default::default()
        }
        .encode_to_vec();
        if with_unknown_group {
            append_key(90, 3, &mut output);
            varint_field(91, 1, &mut output);
            append_key(90, 4, &mut output);
        }
        output
    }

    fn appearance_style_object(
        identifier: u64,
        parent_identifier: Option<u64>,
        overrides: [Option<bool>; 7],
        with_unknown_group: bool,
    ) -> ArchiveObject {
        raw_object(
            identifier,
            vec![(
                TABLE_STYLE_MESSAGE_TYPE,
                appearance_style_payload(
                    identifier,
                    parent_identifier,
                    overrides,
                    with_unknown_group,
                ),
            )],
        )
    }

    fn appearance_preset_payload(network_identifier: u64) -> Vec<u8> {
        tst::TableStylePresetArchive {
            style_network: Some(reference(network_identifier)),
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn appearance_network_payload(table_style_identifier: u64) -> Vec<u8> {
        tst::TableStyleNetworkArchive {
            body_text_style: reference(1),
            header_row_text_style: reference(1),
            header_column_text_style: reference(1),
            footer_row_text_style: reference(1),
            body_cell_style: reference(1),
            header_row_style: reference(1),
            header_column_style: reference(1),
            footer_row_style: reference(1),
            table_style: reference(table_style_identifier),
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn appearance_package(objects: Vec<ArchiveObject>) -> IWorkPackage {
        let mut package = IWorkPackage::new();
        package
            .replace_archive("Index/CalculationEngine.iwa", &Archive { objects })
            .expect("synthetic appearance archive");
        package
    }

    fn appearance_package_in_members(
        calculation_objects: Vec<ArchiveObject>,
        stylesheet_objects: Vec<ArchiveObject>,
    ) -> IWorkPackage {
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/CalculationEngine.iwa",
                &Archive {
                    objects: calculation_objects,
                },
            )
            .expect("synthetic calculation archive");
        package
            .replace_archive(
                "Index/DocumentStylesheet.iwa",
                &Archive {
                    objects: stylesheet_objects,
                },
            )
            .expect("synthetic stylesheet archive");
        package
    }

    fn assert_catalog_appearance_matches_legacy(
        package: &IWorkPackage,
        model_identifier: u64,
    ) -> CommonTableAppearance {
        let before = package.to_bytes().expect("source bytes");
        let expected = crate::table_appearance::table_appearance(package, model_identifier)
            .expect("legacy appearance");
        let mut catalog = KeynoteObjectCatalog::build(package).expect("catalog");
        let facts = catalog_table_model_facts(package, &mut catalog, model_identifier)
            .expect("appearance model facts");
        let actual = catalog_table_appearance(
            package,
            &mut catalog,
            facts.style_identifier,
            facts.style_preset_identifier,
        )
        .expect("catalog appearance");
        assert_eq!(actual, expected);
        assert_eq!(package.to_bytes().expect("source bytes"), before);
        actual
    }

    fn assert_catalog_appearance_rejected(
        package: &IWorkPackage,
        style_identifier: u64,
        style_preset_identifier: Option<u64>,
    ) {
        let before = package.to_bytes().expect("source bytes");
        let mut catalog = KeynoteObjectCatalog::build(package).expect("catalog");
        assert!(
            catalog_table_appearance(
                package,
                &mut catalog,
                style_identifier,
                style_preset_identifier,
            )
            .is_err()
        );
        assert_eq!(package.to_bytes().expect("source bytes"), before);
    }

    fn unmark_source_built_document(editor: KeynoteEditor) -> KeynoteEditor {
        let mut package = editor.into_package();
        package
            .update_archive("Index/Document.iwa", |archive| {
                let object = archive
                    .object_mut(1)
                    .ok_or_else(|| Error::InvalidFormat("document root is missing".to_owned()))?;
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
                    .ok_or_else(|| {
                        Error::InvalidFormat("document payload is missing".to_owned())
                    })?;
                let mut document =
                    kn::DocumentArchive::decode(object.messages[message_index].data.as_slice())?;
                document.super_.template_identifier = Some("Application/Keynote/Native".to_owned());
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: DOCUMENT_MESSAGE_TYPE,
                        data: document.encode_to_vec(),
                    },
                )?;
                Ok(())
            })
            .expect("unmark source-built document");
        KeynoteEditor::from_bytes(&package.to_bytes().expect("unmarked package bytes"))
            .expect("reopen unmarked package")
    }

    #[test]
    fn source_built_reopen_keeps_compatibility_appearance_route() {
        let editor = KeynoteDocumentBuilder::new().build().expect("builder");
        assert!(
            focused_table_appearance_package(editor.package())
                .expect("source-built route")
                .is_none()
        );

        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().expect("builder bytes"))
            .expect("reopen builder");
        assert!(
            focused_table_appearance_package(reopened.package())
                .expect("reopened source-built route")
                .is_none()
        );
    }

    #[test]
    fn exact_unmarked_package_is_not_allowed_to_fallback_after_focused_refusal() {
        let mut editor = KeynoteDocumentBuilder::new().build().expect("builder");
        let geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 40.0, y: 40.0 }),
            size: Some(DrawableSize {
                width: 320.0,
                height: 180.0,
            }),
            flags: Some(3),
            angle: Some(0.0),
        };
        editor
            .add_slide_table(
                0,
                "Focused route",
                2,
                2,
                geometry.position.unwrap(),
                geometry.size.unwrap(),
            )
            .expect("table");
        let editor = unmark_source_built_document(editor);
        assert!(
            focused_table_appearance_package(editor.package())
                .expect("exact route")
                .is_some()
        );

        let focused =
            litchi_keynote::Package::from_bytes(&editor.to_bytes().expect("exact package bytes"))
                .expect("focused package ingress");
        let focused_result = focused.slide_table_appearance(
            litchi_keynote::SlideSelector::index(0),
            litchi_keynote::TableSelector::index(0),
        );
        let host_result = editor.slide_tables(0);
        match (focused_result, host_result) {
            (Ok(expected), Ok(tables)) => {
                assert_eq!(tables[0].appearance, expected);
            },
            (Err(focused_error), Err(host_error)) => {
                let host_error = host_error.to_string();
                assert!(host_error.contains("focused Keynote table appearance"));
                assert!(host_error.contains(&focused_error.to_string()));
            },
            (focused, host) => panic!(
                "focused and compatibility routes diverged: focused={focused:?}, host={host:?}"
            ),
        }
    }

    fn append_varint(value: u64, output: &mut Vec<u8>) {
        let mut value = value;
        while value >= 128 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn append_key(field: u32, wire: u8, output: &mut Vec<u8>) {
        append_varint(u64::from(field) << 3 | u64::from(wire), output);
    }

    fn varint_field(field: u32, value: u64, output: &mut Vec<u8>) {
        append_key(field, 0, output);
        append_varint(value, output);
    }

    fn bytes_field(field: u32, value: &[u8], output: &mut Vec<u8>) {
        append_key(field, 2, output);
        append_varint(value.len() as u64, output);
        output.extend_from_slice(value);
    }

    fn fixed64_field(field: u32, output: &mut Vec<u8>) {
        append_key(field, 1, output);
        output.extend_from_slice(&[0; 8]);
    }

    #[test]
    fn graph_model_discovery_reads_checked_in_native_keynote_models() {
        let fixture = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote/table-discovery.key");
        let source = std::fs::read(fixture).expect("native Keynote fixture");
        let editor = KeynoteEditor::from_bytes(&source).expect("native Keynote editor");
        let before = editor.to_bytes().expect("native package bytes");
        assert_eq!(before, source);
        let listed = editor.slide_tables(0).expect("native table listing");

        assert_eq!(listed.len(), 1);
        let table = &listed[0];
        assert_eq!(table.name, "Table 1");
        assert_eq!((table.rows, table.columns), (5, 4));

        let resolved = slide_table_graph(&editor, 0, table.drawable_object_id)
            .expect("native direct table graph");
        assert_eq!(&resolved.info, table);
        assert_eq!(editor.to_bytes().expect("native package bytes"), source);
    }

    #[test]
    fn direct_table_graph_matches_catalog_listing_without_mutation() {
        let mut editor = KeynoteDocumentBuilder::new().build().expect("builder");
        let geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 40.0, y: 40.0 }),
            size: Some(DrawableSize {
                width: 320.0,
                height: 180.0,
            }),
            flags: Some(3),
            angle: Some(0.0),
        };
        let table = editor
            .add_slide_table(
                0,
                "Catalog graph",
                2,
                2,
                geometry.position.expect("position"),
                geometry.size.expect("size"),
            )
            .expect("table");
        let before = editor.to_bytes().expect("source bytes");
        let listed = editor.slide_tables(0).expect("catalog listing");
        let resolved = slide_table_graph(&editor, 0, table.drawable_object_id)
            .expect("catalog graph resolution");

        assert_eq!(resolved.info, listed[0]);
        assert_eq!(editor.to_bytes().expect("source bytes"), before);
    }

    fn assert_model_rejected_without_fallback(messages: Vec<(u32, Vec<u8>)>) {
        let package = model_package(messages);
        let before = package.to_bytes().expect("package bytes");
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        assert!(catalog_table_model_facts(&package, &mut catalog, 42).is_err());
        assert_eq!(package.to_bytes().expect("package bytes"), before);
    }

    #[test]
    fn canonical_model_precedes_legacy_and_malformed_canonical_never_falls_back() {
        let payload = model_payload();
        let package = model_package(vec![(MODEL_TYPE, payload.clone())]);
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        let facts = catalog_table_model_facts(&package, &mut catalog, 42).expect("canonical");
        assert_eq!((facts.rows, facts.columns), (3, 4));
        assert_eq!(facts.name, "Table");

        assert_model_rejected_without_fallback(vec![
            (MODEL_TYPE, vec![0x08, 0x01]),
            (LEGACY_MODEL_TYPE, payload.clone()),
        ]);
        assert_model_rejected_without_fallback(vec![
            (MODEL_TYPE, payload.clone()),
            (LEGACY_MODEL_TYPE, payload),
        ]);
    }

    #[test]
    fn duplicate_canonical_model_payloads_are_rejected_atomically() {
        let payload = model_payload();
        assert_model_rejected_without_fallback(vec![
            (MODEL_TYPE, payload.clone()),
            (MODEL_TYPE, payload),
        ]);
    }

    #[test]
    fn legacy_model_is_admitted_only_as_the_exact_legacy_candidate() {
        let package = model_package(vec![(LEGACY_MODEL_TYPE, model_payload())]);
        let before = package.to_bytes().expect("package bytes");
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        let facts = catalog_table_model_facts(&package, &mut catalog, 42).expect("legacy");
        assert_eq!((facts.rows, facts.columns), (3, 4));
        assert_eq!(facts.name, "Table");
        assert_eq!(package.to_bytes().expect("package bytes"), before);
    }

    #[test]
    fn table_info_shaped_legacy_candidate_is_not_promoted_as_a_template() {
        let package = table_template_package(vec![(LEGACY_MODEL_TYPE, table_info_payload(42))]);
        let before = package.to_bytes().expect("package bytes");
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        assert!(table_template_from_catalog(&package, &mut catalog).is_err());
        assert_eq!(package.to_bytes().expect("package bytes"), before);
    }

    #[test]
    fn table_info_role_alias_is_not_promoted_as_a_template() {
        let package = table_template_package_with_info_messages(
            vec![
                (TABLE_INFO_MESSAGE_TYPE, table_info_payload(42)),
                (MODEL_TYPE, model_payload()),
            ],
            vec![(MODEL_TYPE, model_payload())],
        );
        let before = package.to_bytes().expect("package bytes");
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        assert!(table_template_from_catalog(&package, &mut catalog).is_err());
        assert_eq!(package.to_bytes().expect("package bytes"), before);
    }

    #[test]
    fn historical_table_model_role_alias_is_rejected() {
        let payload = model_payload();
        for messages in [
            vec![
                (MODEL_TYPE, payload.clone()),
                (TABLE_MODEL_ROLE_ALIAS_MESSAGE_TYPE, payload.clone()),
            ],
            vec![
                (LEGACY_MODEL_TYPE, payload.clone()),
                (TABLE_MODEL_ROLE_ALIAS_MESSAGE_TYPE, payload.clone()),
            ],
            vec![(TABLE_MODEL_ROLE_ALIAS_MESSAGE_TYPE, payload.clone())],
        ] {
            assert_model_rejected_without_fallback(messages);
        }
    }

    #[test]
    fn table_model_appearance_role_aliases_are_rejected_atomically() {
        let payload = model_payload();
        for alias in [
            TABLE_STYLE_MESSAGE_TYPE,
            TABLE_STYLE_PRESET_MESSAGE_TYPE,
            TABLE_STYLE_NETWORK_MESSAGE_TYPE,
        ] {
            assert_model_rejected_without_fallback(vec![
                (MODEL_TYPE, payload.clone()),
                (alias, Vec::new()),
            ]);
        }
    }

    #[test]
    fn historical_table_info_role_alias_is_rejected() {
        let package = table_template_package_with_info_messages(
            vec![
                (TABLE_INFO_MESSAGE_TYPE, table_info_payload(42)),
                (TABLE_MODEL_ROLE_ALIAS_MESSAGE_TYPE, model_payload()),
            ],
            vec![(MODEL_TYPE, model_payload())],
        );
        let before = package.to_bytes().expect("package bytes");
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        assert!(table_template_from_catalog(&package, &mut catalog).is_err());
        assert_eq!(package.to_bytes().expect("package bytes"), before);
    }

    #[test]
    fn multiple_valid_catalog_templates_are_rejected() {
        let payload = model_payload();
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/CalculationEngine.iwa",
                &Archive {
                    objects: vec![
                        raw_object(10, vec![(TABLE_INFO_MESSAGE_TYPE, table_info_payload(42))]),
                        raw_object(11, vec![(TABLE_INFO_MESSAGE_TYPE, table_info_payload(43))]),
                        raw_object(42, vec![(MODEL_TYPE, payload.clone())]),
                        raw_object(43, vec![(MODEL_TYPE, payload)]),
                    ],
                },
            )
            .expect("synthetic archive");
        let before = package.to_bytes().expect("package bytes");
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        assert!(table_template_from_catalog(&package, &mut catalog).is_err());
        assert_eq!(package.to_bytes().expect("package bytes"), before);
    }

    #[test]
    fn catalog_appearance_direct_preset_inherited_and_default_match_legacy() {
        let direct = [
            Some(true),
            Some(true),
            Some(false),
            Some(false),
            Some(true),
            Some(false),
            Some(true),
        ];
        let parent = [
            Some(false),
            Some(true),
            Some(true),
            Some(false),
            Some(false),
            Some(true),
            Some(false),
        ];
        let child = [Some(true), None, None, Some(true), None, None, None];
        let package = appearance_package(vec![
            appearance_model_object(42, 100, None),
            appearance_model_object(43, 0, Some(200)),
            appearance_model_object(44, 102, None),
            appearance_model_object(45, 0, None),
            appearance_style_object(100, None, direct, false),
            appearance_style_object(101, None, parent, false),
            appearance_style_object(102, Some(101), child, false),
            raw_object(
                200,
                vec![(
                    TABLE_STYLE_PRESET_MESSAGE_TYPE,
                    appearance_preset_payload(300),
                )],
            ),
            raw_object(
                300,
                vec![(
                    TABLE_STYLE_NETWORK_MESSAGE_TYPE,
                    appearance_network_payload(102),
                )],
            ),
        ]);
        for model_identifier in [42, 43, 44, 45] {
            assert_catalog_appearance_matches_legacy(&package, model_identifier);
        }
    }

    #[test]
    fn catalog_appearance_accepts_unknown_groups_and_cross_member_style_routes() {
        let package = appearance_package_in_members(
            vec![appearance_model_object(42, 100, None)],
            vec![appearance_style_object(
                100,
                None,
                [
                    Some(true),
                    Some(false),
                    Some(false),
                    Some(true),
                    Some(true),
                    Some(false),
                    Some(true),
                ],
                true,
            )],
        );
        let before = package.to_bytes().expect("source bytes");
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        let initial = catalog.stats();
        let facts =
            catalog_table_model_facts(&package, &mut catalog, 42).expect("appearance model facts");
        let actual = catalog_table_appearance(
            &package,
            &mut catalog,
            facts.style_identifier,
            facts.style_preset_identifier,
        )
        .expect("cross-member appearance");
        assert_eq!(
            actual,
            CommonTableAppearance {
                row_banding: Banding::Enabled,
                row_sizing: RowSizing::Fixed,
                gridlines: Gridlines {
                    body_horizontal: GridlineVisibility::Hidden,
                    header_columns_horizontal: GridlineVisibility::Visible,
                    body_vertical: GridlineVisibility::Visible,
                    header_rows_vertical: GridlineVisibility::Hidden,
                    footer_rows_vertical: GridlineVisibility::Visible,
                },
            }
        );
        assert_eq!(package.to_bytes().expect("source bytes"), before);
        let final_stats = catalog.stats();
        assert_eq!(initial.archives_scanned, final_stats.archives_scanned);
        assert_eq!(initial.archives_scanned, 2);
        assert_eq!(
            final_stats.archive_reads,
            initial.archive_reads.saturating_add(2)
        );
        assert_eq!(
            final_stats.semantic_decodes,
            initial.semantic_decodes.saturating_add(2)
        );
        assert_eq!(final_stats.peak_live_archives, 1);
        assert_eq!(final_stats.retained_payload_bytes, 0);
    }

    #[test]
    fn catalog_appearance_rejects_missing_wrong_duplicate_malformed_and_cycles_atomically() {
        let cases = [
            (appearance_package(vec![]), 100, None, "missing style"),
            (
                appearance_package(vec![raw_object(
                    100,
                    vec![(
                        TABLE_STYLE_PRESET_MESSAGE_TYPE,
                        appearance_preset_payload(300),
                    )],
                )]),
                100,
                None,
                "wrong style type",
            ),
            (
                appearance_package(vec![raw_object(
                    100,
                    vec![
                        (
                            TABLE_STYLE_MESSAGE_TYPE,
                            appearance_style_payload(100, None, [None; 7], false),
                        ),
                        (
                            TABLE_STYLE_MESSAGE_TYPE,
                            appearance_style_payload(100, None, [None; 7], false),
                        ),
                    ],
                )]),
                100,
                None,
                "duplicate style",
            ),
            (
                appearance_package(vec![raw_object(
                    100,
                    vec![(TABLE_STYLE_MESSAGE_TYPE, vec![0x0a])],
                )]),
                100,
                None,
                "malformed style",
            ),
            (
                appearance_package(vec![appearance_style_object(
                    100,
                    Some(101),
                    [None; 7],
                    false,
                )]),
                100,
                None,
                "missing parent",
            ),
            (
                appearance_package(vec![
                    appearance_style_object(100, Some(101), [None; 7], false),
                    appearance_style_object(101, Some(100), [None; 7], false),
                ]),
                100,
                None,
                "style cycle",
            ),
            (
                appearance_package(vec![
                    raw_object(
                        200,
                        vec![(
                            TABLE_STYLE_PRESET_MESSAGE_TYPE,
                            appearance_preset_payload(300),
                        )],
                    ),
                    raw_object(
                        300,
                        vec![(
                            TABLE_STYLE_MESSAGE_TYPE,
                            appearance_style_payload(300, None, [None; 7], false),
                        )],
                    ),
                ]),
                0,
                Some(200),
                "wrong network type",
            ),
            (
                appearance_package(vec![raw_object(
                    200,
                    vec![(
                        TABLE_STYLE_PRESET_MESSAGE_TYPE,
                        appearance_preset_payload(300),
                    )],
                )]),
                0,
                Some(200),
                "missing network",
            ),
            (
                appearance_package(vec![
                    raw_object(
                        200,
                        vec![(
                            TABLE_STYLE_PRESET_MESSAGE_TYPE,
                            appearance_preset_payload(300),
                        )],
                    ),
                    raw_object(300, vec![(TABLE_STYLE_NETWORK_MESSAGE_TYPE, Vec::new())]),
                ]),
                0,
                Some(200),
                "malformed network",
            ),
        ];
        for (package, style_identifier, style_preset_identifier, _label) in cases {
            assert_catalog_appearance_rejected(&package, style_identifier, style_preset_identifier);
        }
    }

    #[test]
    fn catalog_appearance_rejects_style_preset_and_network_role_aliases_atomically() {
        let style_payload = appearance_style_payload(100, None, [None; 7], false);
        let preset_payload = appearance_preset_payload(300);
        let network_payload = appearance_network_payload(100);
        let style_alias = appearance_package(vec![raw_object(
            100,
            vec![
                (TABLE_STYLE_MESSAGE_TYPE, style_payload.clone()),
                (TABLE_STYLE_PRESET_MESSAGE_TYPE, preset_payload.clone()),
            ],
        )]);
        assert_catalog_appearance_rejected(&style_alias, 100, None);
        let preset_alias = appearance_package(vec![
            raw_object(100, vec![(TABLE_STYLE_MESSAGE_TYPE, style_payload.clone())]),
            raw_object(
                200,
                vec![
                    (TABLE_STYLE_PRESET_MESSAGE_TYPE, preset_payload.clone()),
                    (TABLE_STYLE_MESSAGE_TYPE, style_payload.clone()),
                ],
            ),
            raw_object(
                300,
                vec![(TABLE_STYLE_NETWORK_MESSAGE_TYPE, network_payload.clone())],
            ),
        ]);
        assert_catalog_appearance_rejected(&preset_alias, 0, Some(200));
        let network_alias = appearance_package(vec![
            raw_object(100, vec![(TABLE_STYLE_MESSAGE_TYPE, style_payload)]),
            raw_object(200, vec![(TABLE_STYLE_PRESET_MESSAGE_TYPE, preset_payload)]),
            raw_object(
                300,
                vec![
                    (TABLE_STYLE_NETWORK_MESSAGE_TYPE, network_payload),
                    (
                        TABLE_STYLE_PRESET_MESSAGE_TYPE,
                        appearance_preset_payload(300),
                    ),
                ],
            ),
        ]);
        assert_catalog_appearance_rejected(&network_alias, 0, Some(200));
    }

    #[test]
    fn catalog_appearance_validates_missing_and_cyclic_parents_after_full_child_overrides() {
        let complete = [
            Some(true),
            Some(true),
            Some(false),
            Some(false),
            Some(true),
            Some(false),
            Some(true),
        ];
        let missing_parent = appearance_package(vec![appearance_style_object(
            100,
            Some(101),
            complete,
            false,
        )]);
        assert_catalog_appearance_rejected(&missing_parent, 100, None);
        let cyclic_parent = appearance_package(vec![
            appearance_style_object(100, Some(101), complete, false),
            appearance_style_object(101, Some(100), [None; 7], false),
        ]);
        assert_catalog_appearance_rejected(&cyclic_parent, 100, None);
    }

    #[test]
    fn catalog_appearance_direct_style_precedes_malformed_or_missing_preset() {
        let complete = [
            Some(true),
            Some(true),
            Some(false),
            Some(false),
            Some(true),
            Some(false),
            Some(true),
        ];
        let missing_preset = appearance_package(vec![
            appearance_model_object(42, 100, Some(200)),
            appearance_style_object(100, None, complete, false),
        ]);
        assert_catalog_appearance_matches_legacy(&missing_preset, 42);
        let malformed_preset = appearance_package(vec![
            appearance_model_object(42, 100, Some(200)),
            appearance_style_object(100, None, complete, false),
            raw_object(200, vec![(TABLE_STYLE_PRESET_MESSAGE_TYPE, vec![0x1a])]),
        ]);
        assert_catalog_appearance_matches_legacy(&malformed_preset, 42);
    }

    #[test]
    fn catalog_appearance_rejects_inheritance_overdepth_without_mutating_source() {
        let mut objects = Vec::new();
        for identifier in 100..=164 {
            let parent = (identifier < 164).then_some(identifier + 1);
            objects.push(appearance_style_object(
                identifier, parent, [None; 7], false,
            ));
        }
        let package = appearance_package(objects);
        let before = package.to_bytes().expect("source bytes");
        let mut catalog = KeynoteObjectCatalog::build(&package).expect("catalog");
        assert!(catalog_table_appearance(&package, &mut catalog, 100, None).is_err());
        assert_eq!(package.to_bytes().expect("source bytes"), before);
        assert_eq!(catalog.stats().retained_payload_bytes, 0);
    }
}
