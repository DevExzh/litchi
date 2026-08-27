//! Slide ownership and native table graph discovery.

use prost::Message;

use super::super::keynote_object_catalog::{
    KeynoteObjectCatalog, KeynoteObjectCatalogError, map_catalog_error,
};
use super::*;
use crate::protobuf::tst::{TableInfoArchive, TableModelArchive};
use litchi_iwa_common::WireLimits;
use litchi_iwa_protos::{table_info_codec, table_model_discovery_codec};

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
            appearance: crate::table_appearance::table_appearance(package, model_id)?,
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
    if catalog
        .message_type_count(model_id, TABLE_MODEL_ROLE_ALIAS_MESSAGE_TYPE)
        .map_err(map_catalog_error)?
        != 0
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote table model {model_id} contains a historical table-model role alias"
        )));
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
                })
            },
        )
        .map_err(map_catalog_error)
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
        bytes_field(3, &[0x0a, 0x01, 0x08, 0x01], &mut output);
        bytes_field(4, &[0x0a, 0x00], &mut output);
        varint_field(6, 3, &mut output);
        varint_field(7, 4, &mut output);
        bytes_field(8, b"Table", &mut output);
        fixed64_field(16, &mut output);
        fixed64_field(17, &mut output);
        for field in 18..=21 {
            bytes_field(field, &[0x0a, 0x01, 0x08, field as u8 - 17], &mut output);
        }
        for field in 24..=27 {
            bytes_field(field, &[0x0a, 0x01, 0x08, field as u8 - 23], &mut output);
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
}
