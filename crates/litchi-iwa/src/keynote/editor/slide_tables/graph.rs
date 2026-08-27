//! Slide ownership and native table graph discovery.

use prost::Message;

use super::super::keynote_object_catalog::{
    KeynoteObjectCatalog, KeynoteObjectCatalogError, map_catalog_error,
};
use super::*;
use crate::protobuf::tst::{TableInfoArchive, TableModelArchive};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::table::appearance::{
    Appearance as CommonTableAppearance, Banding, GridlineVisibility, Gridlines, RowSizing,
};
use litchi_iwa_protos::{table_appearance_codec, table_info_codec, table_model_discovery_codec};

#[derive(Debug, Clone)]
pub(super) struct SlideTableGraph {
    pub(super) info: KeynoteSlideTableInfo,
    pub(super) table_archive: String,
    pub(super) slide_archive: String,
    pub(super) slide_component_id: u64,
}

#[derive(Debug, Clone)]
pub(super) struct CatalogSlideContext {
    pub(super) slide_id: u64,
    pub(super) slide: kn::SlideArchive,
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

pub(super) fn slide_table_graph(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<SlideTableGraph> {
    let graph = ObjectGraph::read(editor.package())?;
    slide_table_graph_from_graph(editor, &graph, slide_index, drawable_object_id)
}

/// Resolve one table from an already-built object graph.
///
/// Table listing walks every drawable in a slide.  Keeping the graph as an
/// explicit input for that inner loop prevents each table from rebuilding the
/// package-wide graph; callers that only need one table can continue using
/// [`slide_table_graph`].  The graph is still the legacy generated/native
/// representation for now, so this seam also gives the bounded catalog a
/// single replacement point without changing mutation callers.
pub(super) fn slide_table_graph_from_graph(
    editor: &KeynoteEditor,
    graph: &ObjectGraph,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<SlideTableGraph> {
    let context = text_box_create::text_box_context(graph, slide_index)?;
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
    validate_graph_table_info_role(graph, drawable_object_id)?;
    let table_info = graph.decode_type::<TableInfoArchive>(
        drawable_object_id,
        TABLE_INFO_MESSAGE_TYPE,
        "TableInfoArchive",
    )?;
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
    let model_id = table_info.table_model.identifier;
    let model = decode_table_model(graph, model_id)?;
    let table_archive = graph.archive_name(drawable_object_id)?.to_owned();
    let slide_archive = graph.archive_name(context.slide_id)?.to_owned();
    let slide_component_id = component_identifier_for_entry(editor.package(), &slide_archive)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide component {slide_archive} is not registered"
            ))
        })?;
    let lock_state = crate::table_lock::table_lock_state(
        editor.package(),
        &table_archive,
        drawable_object_id,
        "Keynote",
    )?;
    Ok(SlideTableGraph {
        info: KeynoteSlideTableInfo {
            slide_index,
            slide_id: context.slide_id,
            drawable_object_id,
            model_object_id: model_id,
            name: model.table_name,
            rows: model.number_of_rows as usize,
            columns: model.number_of_columns as usize,
            geometry: crate::shapes::geometry_from_drawable(&table_info.super_)?,
            appearance: crate::table_appearance::table_appearance(editor.package(), model_id)?,
            lock_state,
        },
        table_archive,
        slide_archive,
        slide_component_id,
    })
}

/// Resolve the root slide objects through the bounded catalog.
///
/// The returned slide value is deliberately short-lived and owns only the
/// selected root projection.  The catalog retains object slots and message
/// descriptors, not the package's decompressed payloads.
pub(super) fn catalog_slide_context(
    package: &IWorkPackage,
    catalog: &mut KeynoteObjectCatalog,
    slide_index: usize,
) -> Result<CatalogSlideContext> {
    let document: kn::DocumentArchive = catalog
        .decode_type(package, 1, DOCUMENT_MESSAGE_TYPE, "KN.DocumentArchive")
        .map_err(map_catalog_error)?;
    let show: kn::ShowArchive = catalog
        .decode_type(
            package,
            document.show.identifier,
            SHOW_MESSAGE_TYPE,
            "KN.ShowArchive",
        )
        .map_err(map_catalog_error)?;
    let node_reference = show.slide_tree.slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            show.slide_tree.slides.len()
        ))
    })?;
    let node_identifier = node_reference.identifier;
    let node: kn::SlideNodeArchive = catalog
        .decode_type(package, node_identifier, 4, "KN.SlideNodeArchive")
        .map_err(map_catalog_error)?;
    let slide_id = node
        .slide
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide node {node_identifier} has no slide reference"
            ))
        })?
        .identifier;
    let slide: kn::SlideArchive = catalog
        .decode_type(package, slide_id, 5, "KN.SlideArchive")
        .map_err(map_catalog_error)?;
    Ok(CatalogSlideContext { slide_id, slide })
}

/// Resolve one table using an already-decoded catalog slide context.
///
/// Listing passes this context for every drawable on the selected slide.  It
/// is intentionally a separate helper so the root/show/node/slide semantic
/// decodes remain constant when a slide contains many tables.
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
    let table_archive = catalog
        .archive_name(drawable_object_id)
        .map_err(map_catalog_error)?
        .to_owned();
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
    let appearance = catalog_table_appearance(
        package,
        catalog,
        model.style_identifier,
        model.style_preset_identifier,
    )?;
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
        table_archive,
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

fn validate_graph_table_info_role(graph: &ObjectGraph, info_id: u64) -> Result<()> {
    let messages = graph.objects.get(&info_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Keynote table-info object {info_id} is missing"))
    })?;
    let table_info_count = messages
        .iter()
        .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        .count();
    if table_info_count != 1 {
        return Err(Error::InvalidFormat(format!(
            "Keynote table-info object {info_id} must contain exactly one table-info payload"
        )));
    }
    if messages
        .iter()
        .any(|message| matches!(message.type_, 6_001 | TABLE_MODEL_ROLE_ALIAS_MESSAGE_TYPE))
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote table-info object {info_id} contains a table-model role alias"
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

fn table_model_options(source: &[u8]) -> table_model_discovery_codec::DecodeOptions {
    table_model_discovery_codec::DecodeOptions::for_source(source)
}

fn decode_table_model(graph: &ObjectGraph, model_id: u64) -> Result<TableModelArchive> {
    let messages = graph.objects.get(&model_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Keynote table model {model_id} is missing"))
    })?;
    if messages
        .iter()
        .any(|message| message.type_ == TABLE_MODEL_ROLE_ALIAS_MESSAGE_TYPE)
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote table model {model_id} contains a historical table-model role alias"
        )));
    }
    let models = messages
        .iter()
        .filter(|message| TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_))
        .filter_map(|message| TableModelArchive::decode(message.data.as_slice()).ok())
        .collect::<Vec<_>>();
    let [model] = models.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Keynote table model {model_id} must contain exactly one table-model payload"
        )));
    };
    Ok(model.clone())
}

#[allow(dead_code)]
pub(super) fn table_template(package: &IWorkPackage) -> Result<(u64, u64)> {
    let graph = ObjectGraph::read(package)?;
    table_template_from_graph(&graph)
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

/// Find a table template in an already-built graph.
pub(super) fn table_template_from_graph(graph: &ObjectGraph) -> Result<(u64, u64)> {
    let mut candidates = graph.objects.keys().copied().collect::<Vec<_>>();
    candidates.sort_unstable();
    for info_id in candidates {
        let Some(messages) = graph.objects.get(&info_id) else {
            continue;
        };
        if !messages
            .iter()
            .any(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        {
            continue;
        }
        validate_graph_table_info_role(graph, info_id)?;
        for message in messages
            .iter()
            .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        {
            let Ok(info) = TableInfoArchive::decode(message.data.as_slice()) else {
                continue;
            };
            let model_id = info.table_model.identifier;
            if model_id != 0 && decode_table_model(graph, model_id).is_ok() {
                return Ok((info_id, model_id));
            }
        }
    }
    Err(Error::InvalidFormat(
        "Keynote package has no native table creation template".to_owned(),
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject, RawMessage};
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
