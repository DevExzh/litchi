//! The builder's first body table (record 0420): the host builder's table
//! graph, built from scratch on the shared `table_create::fresh` graph.
//!
//! The host Pages builder imports the host Numbers builder's table: the
//! drawable (parented to the body storage), its model and storage, the
//! preset and style network, a stylesheet of its own with the default
//! table styles, the function-browser state, the custom-format list and a
//! calculation engine holding the table's formula owner, under the Numbers
//! builder's identifiers (9 to 41, which the Pages template leaves free).
//! This module writes the same graph into the built-in template, with no
//! bundled table template and no Numbers dependency:
//!
//! - the drawable at the body's start, 468 points wide and 22.73 points per
//!   row, and its attachment (41) anchored by `U+FFFC` at the end of the
//!   body text;
//! - the document root naming the engine, the function-browser state and
//!   the custom-format list, and the theme naming the table preset;
//! - a new `CalculationEngine` component, between the annotation storage
//!   and the document component, as the host registers it;
//! - registry UUIDs for every new object and the document's references to
//!   the new styles and engine.
//!
//! Object UUIDs and the table UUID derive from the document identity (the
//! host draws random ones): the documented UUID policy, so the output is
//! Tier B against the host. Later tables go through
//! [`crate::table::creation`], as the host builder adds them with its
//! editor.
//!
//! The same graph is the first table [`crate::table::creation`] inserts
//! into a table-less built-in-template document at any body position
//! (record 0426; the host editor's `add_table` bootstraps it with
//! `bootstrap_first_table_graph` on such a document and refuses any other
//! table-less one).

use litchi_iwa_archive::package::{EntryEdit, EntryInsertion};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::varint::encode_varint_into;
use litchi_iwa_common::wire::parse_wire_view;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_numbers_wire::table_create::{Frame, OwnerGeometry, Shape, TableUuid, fresh, mix};

use super::Package;
use crate::{PackageError, PackageResult};

pub(crate) const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
const STYLESHEET_COMPONENT: &str = "Index/DocumentStylesheet.iwa";
const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const ENGINE_COMPONENT: &str = "Index/CalculationEngine.iwa";
/// The template's document root, theme, body storage and registry object.
const DOCUMENT: u64 = 1;
const METADATA: u64 = 2;
const THEME: u64 = 102;
const BODY: u64 = 103;
/// The Pages stylesheet component's identifier.
const STYLESHEET_COMPONENT_ID: u64 = 101;
const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;
const THEME_MESSAGE_TYPE: u32 = 10_001;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const FUNCTION_BROWSER_STATE_MESSAGE_TYPE: u32 = 601;
const CUSTOM_FORMAT_LIST_MESSAGE_TYPE: u32 = 222;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;

/// The host builder's identifiers (the Numbers builder's).
const IDS: fresh::Ids = fresh::Ids {
    info: 9,
    model: 10,
    tile: 22,
    row_headers: 23,
    column_headers: 24,
    string_list: 25,
    style_list: 26,
    formula_list: 27,
    format_list: 28,
    control_cell_spec_list: None,
    uid_map: 29,
    stroke_sidecar: 30,
};
const STYLES: fresh::StyleIds = fresh::StyleIds {
    stylesheet: 40,
    list: 11,
    paragraph: 12,
    character: 13,
    shape: 14,
    media: 15,
    drop_cap: 16,
    sheet: 17,
    table: 18,
    cell: 19,
};
const PRESET: u64 = 20;
const ENGINE: u64 = 31;
const NETWORK: u64 = 32;
const FUNCTION_BROWSER_STATE: u64 = 33;
const CUSTOM_FORMAT_LIST: u64 = 34;
const OWNER: u64 = 39;
const ATTACHMENT: u64 = 41;
/// The host Numbers builder's internal identifier of the table's owner.
const OWNER_INTERNAL: u32 = 6;
/// The calculation engine component's versions.
const ENGINE_VERSIONS: [u32; 3] = [3, 2, 10];
/// The host Pages builder's table frame.
const TABLE_WIDTH: f32 = 468.0;
const TABLE_ROW_HEIGHT: f32 = 22.73;

/// The first identifier of the graph (the UUID seed's input), the new
/// drawable, and the objects and components a creation adds or rewrites
/// (the document, stylesheet and registry, and the new engine).
pub(crate) const FIRST_IDENTIFIER: u64 = IDS.info;
pub(crate) const DRAWABLE: u64 = IDS.info;
pub(crate) const CREATED_OBJECTS: usize = DOCUMENT_OBJECTS.len() + STYLE_OBJECTS.len() + 3;
pub(crate) const TOUCHED_COMPONENTS: usize = 4;

/// The document objects the table adds, in the host's order.
const DOCUMENT_OBJECTS: [u64; 15] = [
    9,
    10,
    PRESET,
    22,
    23,
    24,
    25,
    26,
    27,
    28,
    29,
    30,
    NETWORK,
    FUNCTION_BROWSER_STATE,
    CUSTOM_FORMAT_LIST,
];
/// The stylesheet objects the table adds, as the host registers them.
const STYLE_OBJECTS: [u64; 10] = [11, 12, 13, 14, 15, 16, 17, 18, 19, 40];

fn invalid(what: &str) -> PackageError {
    PackageError::InvalidFormat(format!("the built Pages table graph: {what}"))
}

fn wire<E: std::fmt::Debug>(error: E) -> PackageError {
    PackageError::InvalidFormat(format!("the built Pages table graph: {error:?}"))
}

/// A field-by-field protobuf writer.
#[derive(Default)]
struct Writer(Vec<u8>);

impl Writer {
    fn varint(&mut self, number: u32, value: u64) -> &mut Self {
        encode_varint_into(&mut self.0, u64::from(number) << 3);
        encode_varint_into(&mut self.0, value);
        self
    }

    fn bytes(&mut self, number: u32, payload: &[u8]) -> &mut Self {
        encode_varint_into(&mut self.0, (u64::from(number) << 3) | 2);
        encode_varint_into(&mut self.0, payload.len() as u64);
        self.0.extend_from_slice(payload);
        self
    }

    fn reference(&mut self, number: u32, identifier: u64) -> &mut Self {
        let mut reference = Vec::with_capacity(10);
        encode_varint_into(&mut reference, 1 << 3);
        encode_varint_into(&mut reference, identifier);
        self.bytes(number, &reference)
    }

    fn float(&mut self, number: u32, value: f32) -> &mut Self {
        encode_varint_into(&mut self.0, (u64::from(number) << 3) | 5);
        self.0.extend_from_slice(&value.to_le_bytes());
        self
    }

    fn packed(&mut self, number: u32, values: &[u32]) -> &mut Self {
        let mut payload = Vec::new();
        for value in values {
            encode_varint_into(&mut payload, u64::from(*value));
        }
        self.bytes(number, &payload)
    }
}

/// `data` with field `number` (`payload`, length-delimited) placed before
/// the first field numbered above it, as a tag-ordered writer places it.
/// An existing field `number` is refused (the template has none).
fn insert_ordered(data: &[u8], number: u32, payload: &[u8]) -> PackageResult<Vec<u8>> {
    let view = parse_wire_view(data, WireLimits::DEFAULT).map_err(wire)?;
    let mut output = Writer(Vec::with_capacity(data.len() + payload.len() + 8));
    let mut placed = false;
    for field in view.fields() {
        if field.number() == number {
            return Err(invalid("the template already holds the field"));
        }
        if !placed && field.number() > number {
            output.bytes(number, payload);
            placed = true;
        }
        output.0.extend_from_slice(field.raw());
    }
    if !placed {
        output.bytes(number, payload);
    }
    Ok(output.0)
}

/// Rewrite the single length-delimited field `number` of `data`.
fn within(
    data: &[u8],
    number: u32,
    rewrite: impl FnOnce(&[u8]) -> PackageResult<Vec<u8>>,
) -> PackageResult<Vec<u8>> {
    let view = parse_wire_view(data, WireLimits::DEFAULT).map_err(wire)?;
    let mut output = Vec::with_capacity(data.len() + 64);
    let mut rewrite = Some(rewrite);
    for field in view.fields() {
        if field.number() == number && field.wire_type() == 2 {
            let rewrite = rewrite.take().ok_or_else(|| invalid("a repeated field"))?;
            let mut writer = Writer(output);
            writer.bytes(number, &rewrite(field.payload())?);
            output = writer.0;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if rewrite.is_some() {
        return Err(invalid("a missing field"));
    }
    Ok(output)
}

/// A built object with the host's message versions.
fn archive_object(
    identifier: u64,
    message_type: u32,
    payload: Vec<u8>,
    references: Vec<u64>,
    limits: litchi_iwa_core::Limits,
) -> PackageResult<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: payload,
        }],
        limits,
    )
    .map_err(wire)?;
    object.archive_info.message_infos[0].versions = fresh::OBJECT_VERSIONS.to_vec().into();
    object.archive_info.message_infos[0].object_references = references;
    Ok(object)
}

fn from_fresh(
    object: fresh::Object,
    limits: litchi_iwa_core::Limits,
) -> PackageResult<ArchiveObject> {
    archive_object(
        object.identifier,
        object.message_type,
        object.payload,
        object.references,
        limits,
    )
}

/// One registry UUID entry for `identifier`.
fn uuid_entry(seed: u64, identifier: u64) -> Vec<u8> {
    let lower = mix(seed ^ identifier.rotate_left(17));
    let upper = mix(lower ^ 0x6c69_7463_6869_0420);
    let mut uuid = Writer::default();
    uuid.varint(1, lower).varint(2, upper);
    let mut entry = Writer::default();
    entry.varint(1, identifier).bytes(2, &uuid.0);
    entry.0
}

/// One component external reference.
fn external(component: u64, object: Option<u64>) -> Vec<u8> {
    let mut reference = Writer::default();
    reference.varint(1, component);
    if let Some(object) = object {
        reference.varint(2, object);
    }
    reference.0
}

/// A component with `externals` appended to its external references (6)
/// and `uuids` to its object UUID map (11), each after the last entry of
/// its field.
fn extend_component(
    component: &[u8],
    externals: &[Vec<u8>],
    uuids: &[Vec<u8>],
) -> PackageResult<Vec<u8>> {
    let view = parse_wire_view(component, WireLimits::DEFAULT).map_err(wire)?;
    let fields: Vec<_> = view.fields().collect();
    let last = |number: u32| fields.iter().rposition(|field| field.number() == number);
    let first_after = |number: u32| fields.iter().position(|field| field.number() > number);
    // The position after which each addition goes (None: before the first
    // later field, else at the end).
    let place = |number: u32| {
        last(number)
            .map(|index| index + 1)
            .or_else(|| first_after(number))
    };
    let mut output = Writer::default();
    let emit = |output: &mut Writer, number: u32, entries: &[Vec<u8>]| {
        for entry in entries {
            output.bytes(number, entry);
        }
    };
    let (external_at, uuid_at) = (place(6), place(11));
    for (index, field) in fields.iter().enumerate() {
        if external_at == Some(index) {
            emit(&mut output, 6, externals);
        }
        if uuid_at == Some(index) {
            emit(&mut output, 11, uuids);
        }
        output.0.extend_from_slice(field.raw());
    }
    if external_at.is_none_or(|index| index >= fields.len()) {
        emit(&mut output, 6, externals);
    }
    if uuid_at.is_none_or(|index| index >= fields.len()) {
        emit(&mut output, 11, uuids);
    }
    Ok(output.0)
}

/// The registry with the table's component, UUIDs and references. The
/// builder lists the engine component before the document component (the
/// host builder's registry); an insertion appends it after the last
/// component (the host editor's `add_component_registration`).
fn metadata(payload: &[u8], seed: u64, appended: bool) -> PackageResult<Vec<u8>> {
    let uuids = |identifiers: &[u64]| -> Vec<Vec<u8>> {
        identifiers
            .iter()
            .map(|identifier| uuid_entry(seed, *identifier))
            .collect()
    };
    let mut engine = Writer::default();
    engine
        .varint(1, ENGINE)
        .bytes(2, b"CalculationEngine")
        .packed(4, &ENGINE_VERSIONS)
        .packed(5, &ENGINE_VERSIONS)
        .bytes(6, &external(DOCUMENT, Some(IDS.info)));
    for entry in uuids(&[ENGINE, OWNER]) {
        engine.bytes(11, &entry);
    }
    engine.varint(12, 1).packed(14, &ENGINE_VERSIONS);
    let mut document_objects = DOCUMENT_OBJECTS.to_vec();
    document_objects.push(ATTACHMENT);
    let mut document_externals: Vec<Vec<u8>> = STYLE_OBJECTS
        .iter()
        .map(|object| external(STYLESHEET_COMPONENT_ID, Some(*object)))
        .collect();
    document_externals.push(external(ENGINE, None));

    let view = parse_wire_view(payload, WireLimits::DEFAULT).map_err(wire)?;
    let last_component = view
        .fields()
        .enumerate()
        .filter(|(_, field)| field.number() == 3)
        .map(|(index, _)| index)
        .last();
    let mut output = Writer::default();
    let mut seen = (false, false);
    for (index, field) in view.fields().enumerate() {
        if field.number() != 3 {
            output.0.extend_from_slice(field.raw());
            continue;
        }
        match identifier_of(field.payload())? {
            Some(DOCUMENT) => {
                if !appended {
                    output.bytes(3, &engine.0);
                }
                output.bytes(
                    3,
                    &extend_component(
                        field.payload(),
                        &document_externals,
                        &uuids(&document_objects),
                    )?,
                );
                seen.0 = true;
            },
            Some(STYLESHEET_COMPONENT_ID) => {
                output.bytes(
                    3,
                    &extend_component(field.payload(), &[], &uuids(&STYLE_OBJECTS))?,
                );
                seen.1 = true;
            },
            _ => {
                output.0.extend_from_slice(field.raw());
            },
        }
        if appended && Some(index) == last_component {
            output.bytes(3, &engine.0);
        }
    }
    if seen != (true, true) {
        return Err(invalid("the template registry"));
    }
    Ok(output.0)
}

/// The identifier (field 1) of a registry component entry.
fn identifier_of(component: &[u8]) -> PackageResult<Option<u64>> {
    Ok(
        litchi_iwa_common::wire::repeated_varint_values(component, 1)
            .map_err(wire)?
            .first()
            .copied(),
    )
}

/// The single message of `message_type` of object `identifier`.
fn message_index(object: &ArchiveObject, message_type: u32) -> PackageResult<usize> {
    match object.messages.as_slice() {
        [message] if message.type_ == message_type => Ok(0),
        _ => Err(invalid("an unexpected template object")),
    }
}

fn replace(
    archive: &mut Archive,
    identifier: u64,
    message_type: u32,
    rewrite: impl FnOnce(&[u8]) -> PackageResult<Vec<u8>>,
    references: &[u64],
    limits: litchi_iwa_core::Limits,
) -> PackageResult<()> {
    let object = archive
        .object_mut(identifier)
        .ok_or_else(|| invalid("a missing template object"))?;
    let index = message_index(object, message_type)?;
    let data = rewrite(&object.messages[index].data)?;
    object
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(wire)?;
    let declared = &mut object.archive_info.message_infos[index].object_references;
    for reference in references {
        if !declared.contains(reference) {
            declared.push(*reference);
        }
    }
    Ok(())
}

/// Whether `package` is a table-less built-in-template document that can
/// take the first table graph (the host's `require_scratch_table_slots`):
/// no calculation engine, the template's body storage, document and
/// stylesheet components, and every identifier of the graph free.
pub(crate) fn admits(package: &Package) -> bool {
    scratch_slots(package).is_ok()
}

fn scratch_slots(package: &Package) -> PackageResult<()> {
    let source = &package.state.source;
    if source.components().get(ENGINE_COMPONENT).is_some() {
        return Err(invalid("the template already has a calculation engine"));
    }
    for member in [DOCUMENT_COMPONENT, STYLESHEET_COMPONENT, METADATA_COMPONENT] {
        if source.components().get(member).is_none() {
            return Err(invalid("a missing template component"));
        }
    }
    let (body, _) = super::body_storage_objects::body_storage(package).map_err(wire)?;
    if body.identifier.get() != BODY {
        return Err(invalid("the template body"));
    }
    let reserved: Vec<u64> = DOCUMENT_OBJECTS
        .iter()
        .chain(&STYLE_OBJECTS)
        .chain(&[ENGINE, OWNER, ATTACHMENT])
        .copied()
        .collect();
    for component in source.components().iter() {
        if component.archive().objects.iter().any(|object| {
            object
                .archive_info
                .identifier
                .is_some_and(|identifier| reserved.contains(&identifier))
        }) {
            return Err(invalid("the template uses the table's identifiers"));
        }
    }
    Ok(())
}

/// `package` (the built-in template with its identity, language, locale and
/// body text written) with the host builder's first table, `name` of
/// `rows` × `columns`, anchored at UTF-16 body position `anchor` (`None`:
/// the end of the body text). `seed` derives the table and object UUIDs.
pub(crate) fn install(
    package: &Package,
    (name, rows, columns): (&str, u32, u32),
    locale: &str,
    seed: u64,
    anchor: Option<usize>,
) -> PackageResult<Package> {
    let source = &package.state.source;
    let limits = source.limits();
    let archive_limits = limits.effective_archive_limits().map_err(wire)?;
    let component = |name: &str| -> PackageResult<Archive> {
        source
            .components()
            .get(name)
            .map(|component| component.archive().clone())
            .ok_or_else(|| invalid("a missing template component"))
    };
    scratch_slots(package)?;
    let (body, _) = super::body_storage_objects::body_storage(package).map_err(wire)?;
    let mut document = component(DOCUMENT_COMPONENT)?;
    let mut stylesheet = component(STYLESHEET_COMPONENT)?;
    let mut registry = component(METADATA_COMPONENT)?;

    // The table graph, its preset and network, and the default styles.
    let uuid = TableUuid::derive(seed);
    let shape = Shape {
        name,
        rows,
        columns,
        row_height: None,
        column_width: None,
    };
    #[expect(
        clippy::cast_precision_loss,
        reason = "the host computes the frame height in single precision"
    )]
    let height = TABLE_ROW_HEIGHT * rows as f32;
    let styles = STYLES.context(PRESET);
    let mut table = fresh::table(
        &IDS,
        fresh::Placement {
            parent: BODY,
            frame: Frame {
                x: 0.0,
                y: 0.0,
                width: TABLE_WIDTH,
                height,
            },
        },
        &styles,
        &shape,
        &uuid.text,
        IDS.model,
        fresh::Profile::Builder,
    )
    .map_err(wire)?;
    // The host keeps the Numbers drawable's model reference and appends
    // the body.
    table[0].references = vec![IDS.model, BODY];
    let geometry = OwnerGeometry::of(&table[1].payload, rows, columns).map_err(wire)?;
    let mut attachment = Writer::default();
    attachment
        .reference(1, IDS.info)
        .varint(2, 0)
        .float(3, 0.0)
        .varint(4, 0)
        .float(5, 0.0);
    let mut browser = Writer::default();
    browser.varint(4, 0);

    // Document: the attachment, then the table's objects in the host's
    // order.
    document
        .insert_object(
            archive_object(
                ATTACHMENT,
                ATTACHMENT_MESSAGE_TYPE,
                attachment.0,
                vec![IDS.info],
                archive_limits,
            )?,
            archive_limits,
        )
        .map_err(wire)?;
    let mut storage = table.into_iter();
    let mut take = |document: &mut Archive| -> PackageResult<()> {
        let object = storage.next().ok_or_else(|| invalid("the table graph"))?;
        document
            .insert_object(from_fresh(object, archive_limits)?, archive_limits)
            .map_err(wire)
    };
    take(&mut document)?;
    take(&mut document)?;
    document
        .insert_object(
            from_fresh(fresh::preset(PRESET, 0, NETWORK), archive_limits)?,
            archive_limits,
        )
        .map_err(wire)?;
    for _ in 2..IDS.count() {
        take(&mut document)?;
    }
    for object in [
        fresh::style_network(NETWORK, &styles),
        fresh::Object {
            identifier: FUNCTION_BROWSER_STATE,
            message_type: FUNCTION_BROWSER_STATE_MESSAGE_TYPE,
            payload: browser.0,
            references: Vec::new(),
        },
        fresh::Object {
            identifier: CUSTOM_FORMAT_LIST,
            message_type: CUSTOM_FORMAT_LIST_MESSAGE_TYPE,
            payload: Vec::new(),
            references: Vec::new(),
        },
    ] {
        document
            .insert_object(from_fresh(object, archive_limits)?, archive_limits)
            .map_err(wire)?;
    }
    for object in fresh::default_styles(&STYLES) {
        stylesheet
            .insert_object(from_fresh(object, archive_limits)?, archive_limits)
            .map_err(wire)?;
    }

    // The document root, the theme and the body story.
    replace(
        &mut document,
        DOCUMENT,
        DOCUMENT_MESSAGE_TYPE,
        |data| {
            // `TP.DocumentArchive.super` (15), the `TSA.DocumentArchive`.
            within(data, 15, |application| {
                let mut application = insert_ordered(
                    application,
                    4,
                    &litchi_numbers_wire::table_create::reference_payload(ENGINE),
                )?;
                application = insert_ordered(
                    &application,
                    6,
                    &litchi_numbers_wire::table_create::reference_payload(FUNCTION_BROWSER_STATE),
                )?;
                insert_ordered(
                    &application,
                    12,
                    &litchi_numbers_wire::table_create::reference_payload(CUSTOM_FORMAT_LIST),
                )
            })
        },
        &[ENGINE, FUNCTION_BROWSER_STATE, CUSTOM_FORMAT_LIST],
        archive_limits,
    )?;
    replace(
        &mut document,
        THEME,
        THEME_MESSAGE_TYPE,
        |data| {
            within(data, 1, |theme| {
                let mut presets = Writer::default();
                presets.reference(1, PRESET);
                insert_ordered(theme, 200, &presets.0)
            })
        },
        &[PRESET],
        archive_limits,
    )?;
    let body_object = document
        .object(BODY)
        .ok_or_else(|| invalid("the template body"))?;
    let body_payload = body_object
        .messages
        .get(body.message_index)
        .ok_or_else(|| invalid("the template body"))?;
    let body_type = body_payload.type_;
    let end = super::body_table_creation::prepared_len(&body_payload.data).map_err(wire)?;
    let inserted = anchor.is_some();
    let anchor = anchor.unwrap_or(end);
    if anchor > end {
        return Err(invalid("the anchor is past the body text"));
    }
    let anchored =
        super::body_table_creation::anchored_body(package, &body_payload.data, anchor, ATTACHMENT)
            .map_err(wire)?;
    replace(
        &mut document,
        BODY,
        body_type,
        |_| Ok(anchored),
        &[ATTACHMENT],
        archive_limits,
    )?;

    // The calculation engine and the table's formula owner.
    let mut engine = Archive {
        objects: Vec::new(),
    };
    let owner_uid = uuid.owner_uid();
    engine
        .insert_object(
            from_fresh(
                fresh::engine(ENGINE, OWNER, OWNER_INTERNAL, owner_uid, locale),
                archive_limits,
            )?,
            archive_limits,
        )
        .map_err(wire)?;
    engine
        .insert_object(
            archive_object(
                OWNER,
                FORMULA_OWNER_MESSAGE_TYPE,
                litchi_numbers_wire::table_create::owner_payload(
                    owner_uid,
                    OWNER_INTERNAL,
                    IDS.info,
                    geometry,
                )
                .map_err(wire)?,
                Vec::new(),
                archive_limits,
            )?,
            archive_limits,
        )
        .map_err(wire)?;

    // The registry.
    replace(
        &mut registry,
        METADATA,
        METADATA_MESSAGE_TYPE,
        |data| metadata(data, seed, inserted),
        &[],
        archive_limits,
    )?;

    // One publication: three rewritten members and the new engine after
    // the registry, where the host's package lists it.
    let snappy = limits.snappy_limits().map_err(wire)?;
    let encode = |archive: &Archive| -> PackageResult<Vec<u8>> {
        let bytes = archive.to_bytes(archive_limits).map_err(wire)?;
        SnappyStream::compress(&bytes, snappy).map_err(wire)
    };
    let (document, stylesheet, registry, engine) = (
        encode(&document)?,
        encode(&stylesheet)?,
        encode(&registry)?,
        encode(&engine)?,
    );
    // The builder lists the engine after the registry member; the host
    // editor's insertion appends it to the package.
    let position = if inserted {
        source.package().len()
    } else {
        source
            .package()
            .iter()
            .position(|entry| entry.name() == METADATA_COMPONENT)
            .ok_or_else(|| invalid("the template registry member"))?
            + 1
    };
    let bytes = source
        .package()
        .reassemble_with_changes_to_bytes(
            &[EntryInsertion::at(position, ENGINE_COMPONENT, &engine)],
            &[
                EntryEdit::new(DOCUMENT_COMPONENT, &document),
                EntryEdit::new(STYLESHEET_COMPONENT, &stylesheet),
                EntryEdit::new(METADATA_COMPONENT, &registry),
            ],
            &[],
            limits,
        )
        .map_err(wire)?;
    Package::from_bytes_with_archive_limits(&bytes, limits)
}
