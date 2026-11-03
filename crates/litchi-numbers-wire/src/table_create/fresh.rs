//! A from-scratch table graph: every object of a new, empty table built from
//! typed values, with no template table to clone (record 0420).
//!
//! This reproduces the table construction without a template of the
//! retired migration host (deleted by record 0429; last present at commit
//! `27b4840684`), in two profiles ([`Profile`]):
//!
//! - [`Profile::Builder`], the table of that host's Numbers builder, which
//!   its Pages builder imported for `body_table`: four data lists and the
//!   model's references in the builder's order;
//! - [`Profile::Bootstrap`], that host editor's table for a workbook that had
//!   no table left to clone (`bootstrap_empty_table_graph`): a fifth,
//!   control-cell-spec list, the BNC format-table reference and the model's
//!   references sorted.
//!
//! The builder is leaf-neutral, as [`super`]'s template clone and the
//! shared `structure::duplicate` are. The caller supplies:
//!
//! - the identifiers ([`Ids`]), allocated under its own watermark policy;
//! - the container: the drawable's parent and frame ([`Placement`]). The
//!   caller attaches the drawable (a sheet's drawable list, a body
//!   attachment and its `U+FFFC`, a slide's drawables);
//! - the stylesheet context ([`Styles`]): the preset and the styles the
//!   model and its style network name, either existing ones (a theme's
//!   first table preset) or a fresh set from [`default_styles`];
//! - the table's identity (its UUID text and the UID-map seed).
//!
//! It returns each object's identifier, message type, payload and declared
//! references ([`Object`]); the caller places them in its components and
//! records their registry consequences. Every payload is encoded field by
//! field in tag order, as the host's Prost writer emits it, so a template
//! the host builder wrote is reproduced byte for byte. No clock or
//! randomness is consulted.

use super::{Frame, Shape, admits_size, reference_payload};
use crate::table_write::WriteError;

type Result<T> = std::result::Result<T, WriteError>;

/// `TST.TableInfoArchive`.
pub const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
/// `TST.TableModelArchive`.
pub const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
/// `TST.Tile`.
pub const TILE_MESSAGE_TYPE: u32 = 6_002;
/// `TST.TableStyleArchive`.
pub const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
/// `TST.CellStyleArchive`.
pub const CELL_STYLE_MESSAGE_TYPE: u32 = 6_004;
/// `TST.TableDataList`.
pub const DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
/// `TST.HeaderStorageBucket`.
pub const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
/// `TST.TableStylePresetArchive`.
pub const PRESET_MESSAGE_TYPE: u32 = 6_008;
/// `TST.TableStyleNetworkArchive`.
pub const STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
/// `TST.ColumnRowUIDMapArchive`.
pub const UID_MAP_MESSAGE_TYPE: u32 = 6_267;
/// `TST.StrokeSidecarArchive`.
pub const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
/// `TSS.StylesheetArchive`.
pub const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// `TSWP.CharacterStyleArchive`.
pub const CHARACTER_STYLE_MESSAGE_TYPE: u32 = 2_021;
/// `TSWP.ParagraphStyleArchive`.
pub const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
/// `TSWP.ListStyleArchive`.
pub const LIST_STYLE_MESSAGE_TYPE: u32 = 2_023;
/// `TSWP.ShapeStyleArchive`.
pub const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
/// `TSD.MediaStyleArchive`.
pub const MEDIA_STYLE_MESSAGE_TYPE: u32 = 3_016;
/// `TSWP.DropCapStyleArchive`.
pub const DROP_CAP_STYLE_MESSAGE_TYPE: u32 = 10_024;
/// `TN.SheetStyleArchive`.
pub const SHEET_STYLE_MESSAGE_TYPE: u32 = 12_050;
/// `TSCE.CalculationEngineArchive`.
pub const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
/// The message versions the host writes on every object it builds.
pub const OBJECT_VERSIONS: [u32; 3] = [1, 0, 5];

/// The host's default drawable frame of a table it builds.
pub const DEFAULT_FRAME: Frame = Frame {
    x: 0.0,
    y: 0.0,
    width: 490.0,
    height: 200.0,
};
/// The host's default row height, in points.
pub const DEFAULT_ROW_HEIGHT: f64 = 20.0;
/// The host's default column width, in points.
pub const DEFAULT_COLUMN_WIDTH: f64 = 98.0;
/// The tile size the host writes.
const TILE_SIZE: u64 = 256;
/// `TableDataList.ListType` values.
const LIST_STRING: u64 = 1;
const LIST_FORMAT: u64 = 2;
const LIST_FORMULA: u64 = 3;
const LIST_STYLE: u64 = 4;
const LIST_CONTROL_CELL_SPEC: u64 = 12;

/// Which host construction the graph follows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Profile {
    /// The host builders' table (Numbers `Builder`, Pages `body_table`).
    Builder,
    /// The host editor's table for a workbook without a table to clone.
    Bootstrap,
}

/// The identifiers of a new table's objects.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Ids {
    /// The drawable (`TST.TableInfoArchive`).
    pub info: u64,
    /// The model.
    pub model: u64,
    /// The one tile.
    pub tile: u64,
    /// The row-header bucket.
    pub row_headers: u64,
    /// The column-header bucket.
    pub column_headers: u64,
    /// The string list.
    pub string_list: u64,
    /// The style list.
    pub style_list: u64,
    /// The formula list.
    pub formula_list: u64,
    /// The format list.
    pub format_list: u64,
    /// The control-cell-spec list ([`Profile::Bootstrap`] only).
    pub control_cell_spec_list: Option<u64>,
    /// The row/column UID map.
    pub uid_map: u64,
    /// The stroke sidecar.
    pub stroke_sidecar: u64,
}

impl Ids {
    /// Consecutive identifiers from `first`, in the host editor's
    /// allocation order (drawable, model, tile, headers, lists, UID map,
    /// stroke sidecar); the control-cell-spec list follows the format list
    /// under [`Profile::Bootstrap`].
    #[must_use]
    pub const fn consecutive(first: u64, profile: Profile) -> Self {
        let control = matches!(profile, Profile::Bootstrap);
        let after_lists = if control { first + 10 } else { first + 9 };
        Self {
            info: first,
            model: first + 1,
            tile: first + 2,
            row_headers: first + 3,
            column_headers: first + 4,
            string_list: first + 5,
            style_list: first + 6,
            formula_list: first + 7,
            format_list: first + 8,
            control_cell_spec_list: if control { Some(first + 9) } else { None },
            uid_map: after_lists,
            stroke_sidecar: after_lists + 1,
        }
    }

    /// How many identifiers the graph takes.
    #[must_use]
    pub const fn count(&self) -> usize {
        if self.control_cell_spec_list.is_some() {
            12
        } else {
            11
        }
    }

    /// Every identifier, in allocation order.
    #[must_use]
    pub fn all(&self) -> Vec<u64> {
        let mut output = vec![
            self.info,
            self.model,
            self.tile,
            self.row_headers,
            self.column_headers,
            self.string_list,
            self.style_list,
            self.formula_list,
            self.format_list,
        ];
        output.extend(self.control_cell_spec_list);
        output.extend([self.uid_map, self.stroke_sidecar]);
        output
    }
}

/// The stylesheet context: the preset and the styles a new table names.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Styles {
    /// The table style preset.
    pub preset: u64,
    /// The model's `preset_index`.
    pub preset_index: u32,
    /// The table style.
    pub table: u64,
    /// The body text style.
    pub body_text: u64,
    /// The header-row text style.
    pub header_row_text: u64,
    /// The header-column text style.
    pub header_column_text: u64,
    /// The footer-row text style.
    pub footer_row_text: u64,
    /// The body cell style.
    pub body_cell: u64,
    /// The header-row cell style.
    pub header_row_cell: u64,
    /// The header-column cell style.
    pub header_column_cell: u64,
    /// The footer-row cell style.
    pub footer_row_cell: u64,
    /// The table-name paragraph style.
    pub table_name: Option<u64>,
    /// The table-name shape style.
    pub table_name_shape: Option<u64>,
}

impl Styles {
    /// One text style and one cell style for every region, as the host
    /// builders write them.
    #[must_use]
    pub const fn uniform(preset: u64, table: u64, text: u64, cell: u64, name_shape: u64) -> Self {
        Self {
            preset,
            preset_index: 0,
            table,
            body_text: text,
            header_row_text: text,
            header_column_text: text,
            footer_row_text: text,
            body_cell: cell,
            header_row_cell: cell,
            header_column_cell: cell,
            footer_row_cell: cell,
            table_name: Some(text),
            table_name_shape: Some(name_shape),
        }
    }

    /// The styles a model names, in the host builder's declaration order
    /// (the table style, the text styles, the name styles, the cell
    /// styles), each once.
    fn model_styles(&self) -> Vec<u64> {
        let mut output = Vec::with_capacity(12);
        let candidates = [
            Some(self.table),
            Some(self.body_text),
            Some(self.header_row_text),
            Some(self.header_column_text),
            Some(self.footer_row_text),
            self.table_name,
            self.table_name_shape,
            Some(self.body_cell),
            Some(self.header_row_cell),
            Some(self.header_column_cell),
            Some(self.footer_row_cell),
        ];
        for identifier in candidates.into_iter().flatten() {
            if !output.contains(&identifier) {
                output.push(identifier);
            }
        }
        output
    }

    /// Every object this context names (its preset first), each once.
    #[must_use]
    pub fn objects(&self) -> Vec<u64> {
        let mut output = vec![self.preset];
        for identifier in self.model_styles() {
            if !output.contains(&identifier) {
                output.push(identifier);
            }
        }
        output
    }
}

/// Where the new drawable goes: its parent container and its frame.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Placement {
    /// The container the drawable names as its parent (a sheet, a body
    /// storage, a slide).
    pub parent: u64,
    /// The drawable's frame.
    pub frame: Frame,
}

/// One built object: its identifier, message type, payload and the
/// references its header declares.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Object {
    /// The object identifier.
    pub identifier: u64,
    /// The message type of its single payload.
    pub message_type: u32,
    /// The payload.
    pub payload: Vec<u8>,
    /// The references the object header declares.
    pub references: Vec<u64>,
}

/// A field-by-field protobuf writer, in the order the caller writes.
#[derive(Default)]
struct Writer(Vec<u8>);

impl Writer {
    fn key(&mut self, number: u32, wire_type: u8) {
        litchi_iwa_common::varint::encode_varint_into(
            &mut self.0,
            (u64::from(number) << 3) | u64::from(wire_type),
        );
    }

    fn varint(&mut self, number: u32, value: u64) -> &mut Self {
        self.key(number, 0);
        litchi_iwa_common::varint::encode_varint_into(&mut self.0, value);
        self
    }

    fn bool(&mut self, number: u32, value: bool) -> &mut Self {
        self.varint(number, u64::from(value))
    }

    fn bytes(&mut self, number: u32, payload: &[u8]) -> &mut Self {
        self.key(number, 2);
        litchi_iwa_common::varint::encode_varint_into(&mut self.0, payload.len() as u64);
        self.0.extend_from_slice(payload);
        self
    }

    fn reference(&mut self, number: u32, identifier: u64) -> &mut Self {
        self.bytes(number, &reference_payload(identifier))
    }

    fn float(&mut self, number: u32, value: f32) -> &mut Self {
        self.key(number, 5);
        self.0.extend_from_slice(&value.to_le_bytes());
        self
    }

    fn double(&mut self, number: u32, value: f64) -> &mut Self {
        self.key(number, 1);
        self.0.extend_from_slice(&value.to_le_bytes());
        self
    }

    fn message(&mut self, number: u32, build: impl FnOnce(&mut Self)) -> &mut Self {
        let mut inner = Self::default();
        build(&mut inner);
        self.bytes(number, &inner.0)
    }

    fn finish(self) -> Vec<u8> {
        self.0
    }
}

fn object(identifier: u64, message_type: u32, payload: Vec<u8>, references: Vec<u64>) -> Object {
    Object {
        identifier,
        message_type,
        payload,
        references,
    }
}

fn finite(value: f32) -> Result<f32> {
    if value.is_finite() {
        Ok(value)
    } else {
        Err(WriteError::Invalid)
    }
}

/// The drawable payload: a `TST.TableInfoArchive` parented to
/// `placement.parent`, naming `model`, at `placement.frame`.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a non-finite or empty frame.
pub fn info_payload(placement: Placement, model: u64) -> Result<Vec<u8>> {
    let frame = placement.frame;
    let (x, y) = (finite(frame.x)?, finite(frame.y)?);
    let (width, height) = (finite(frame.width)?, finite(frame.height)?);
    if width <= 0.0 || height <= 0.0 {
        return Err(WriteError::Invalid);
    }
    let mut info = Writer::default();
    info.message(1, |drawable| {
        drawable
            .message(1, |geometry| {
                geometry
                    .message(1, |point| {
                        point.float(1, x).float(2, y);
                    })
                    .message(2, |size| {
                        size.float(1, width).float(2, height);
                    })
                    .varint(3, 3)
                    .float(4, 0.0);
            })
            .reference(2, placement.parent)
            .bool(5, false)
            .bool(7, false)
            .bool(12, true)
            .bool(13, true);
    })
    .reference(2, model)
    .varint(10, 0);
    Ok(info.finish())
}

/// The model payload.
fn model_payload(
    ids: &Ids,
    styles: &Styles,
    shape: &Shape<'_>,
    uuid: &str,
    profile: Profile,
) -> Result<Vec<u8>> {
    let row_height = shape.row_height.unwrap_or(DEFAULT_ROW_HEIGHT);
    let column_width = shape.column_width.unwrap_or(DEFAULT_COLUMN_WIDTH);
    if !(row_height.is_finite() && column_width.is_finite())
        || row_height <= 0.0
        || column_width <= 0.0
    {
        return Err(WriteError::Invalid);
    }
    let mut model = Writer::default();
    model
        .bytes(1, uuid.as_bytes())
        .reference(3, styles.table)
        .message(4, |store| {
            store
                .message(1, |headers| {
                    headers.varint(1, 1).reference(2, ids.row_headers);
                })
                .reference(2, ids.column_headers)
                .message(3, |tiles| {
                    tiles
                        .message(1, |tile| {
                            tile.varint(1, 0).reference(2, ids.tile);
                        })
                        .varint(2, TILE_SIZE);
                })
                .reference(4, ids.string_list)
                .reference(5, ids.style_list)
                .reference(6, ids.formula_list)
                .varint(7, 1)
                .varint(8, 0)
                .message(9, |tree| {
                    tree.message(1, |node| {
                        node.varint(1, 0).varint(2, 0);
                    });
                })
                .bytes(10, &[])
                .reference(11, ids.format_list)
                .varint(14, 4);
            if let Some(control) = ids.control_cell_spec_list {
                store.reference(21, control).reference(22, ids.format_list);
            }
        })
        .varint(6, u64::from(shape.rows))
        .varint(7, u64::from(shape.columns))
        .bytes(8, shape.name.as_bytes())
        .varint(9, 1)
        .varint(10, 1)
        .bool(12, true)
        .bool(13, true)
        .double(16, row_height)
        .double(17, column_width)
        .reference(18, styles.body_cell)
        .reference(19, styles.header_row_cell)
        .reference(20, styles.header_column_cell)
        .reference(21, styles.footer_row_cell)
        .bool(22, false)
        .reference(24, styles.body_text)
        .reference(25, styles.header_row_text)
        .reference(26, styles.header_column_text)
        .reference(27, styles.footer_row_text)
        .varint(28, u64::from(styles.preset_index))
        .bool(29, true);
    if let Some(name) = styles.table_name {
        model.reference(30, name);
    }
    model.bool(31, false).bool(32, true);
    // The bootstrap writes the title height only when both name styles are
    // named; the builders always write it.
    let named = styles.table_name.is_some() && styles.table_name_shape.is_some();
    if matches!(profile, Profile::Builder) || named {
        model.double(33, 0.0);
    }
    if let Some(shape_style) = styles.table_name_shape {
        model.reference(36, shape_style);
    }
    model
        .reference(46, ids.uid_map)
        .reference(48, styles.preset)
        .reference(49, ids.stroke_sidecar);
    Ok(model.finish())
}

/// The references a model declares: its styles and preset, then its
/// storage (sorted under [`Profile::Bootstrap`], as the host editor lists
/// them).
fn model_references(ids: &Ids, styles: &Styles, profile: Profile) -> Vec<u64> {
    let mut output = styles.model_styles();
    for identifier in std::iter::once(styles.preset).chain(ids.all().into_iter().skip(2)) {
        if !output.contains(&identifier) {
            output.push(identifier);
        }
    }
    if matches!(profile, Profile::Bootstrap) {
        output.sort_unstable();
    }
    output
}

fn data_list(list_type: u64) -> Vec<u8> {
    let mut list = Writer::default();
    list.varint(1, list_type).varint(2, 1).bool(5, true);
    list.finish()
}

/// Every object of a new, empty `shape.rows` × `shape.columns` table, in
/// the host's order: the drawable, the model, the tile, the header buckets,
/// the data lists, the UID map and the stroke sidecar. The drawable
/// declares `[parent, model]`; `uuid` is the model's `table_id`, and
/// `uid_seed` derives the UID map's row and column identities (the shared
/// fresh-map identities of [`super::empty_storage`]).
///
/// # Errors
///
/// [`WriteError::Invalid`] for a size [`admits_size`] refuses, a profile
/// and identifier set that disagree (a control-cell-spec list only under
/// [`Profile::Bootstrap`]), or an invalid frame or default size.
pub fn table(
    ids: &Ids,
    placement: Placement,
    styles: &Styles,
    shape: &Shape<'_>,
    uuid: &str,
    uid_seed: u64,
    profile: Profile,
) -> Result<Vec<Object>> {
    if !admits_size(shape.rows, shape.columns)
        || ids.control_cell_spec_list.is_some() != matches!(profile, Profile::Bootstrap)
    {
        return Err(WriteError::Invalid);
    }
    let mut objects = Vec::with_capacity(ids.count());
    objects.push(object(
        ids.info,
        TABLE_INFO_MESSAGE_TYPE,
        info_payload(placement, ids.model)?,
        vec![placement.parent, ids.model],
    ));
    objects.push(object(
        ids.model,
        TABLE_MODEL_MESSAGE_TYPE,
        model_payload(ids, styles, shape, uuid, profile)?,
        model_references(ids, styles, profile),
    ));
    let mut tile = Writer::default();
    tile.varint(1, u64::from(shape.columns - 1))
        .varint(2, u64::from(shape.rows - 1))
        .varint(3, 0)
        .varint(4, 0)
        .varint(6, 5)
        .bool(7, true);
    objects.push(object(
        ids.tile,
        TILE_MESSAGE_TYPE,
        tile.finish(),
        Vec::new(),
    ));
    for identifier in [ids.row_headers, ids.column_headers] {
        let mut bucket = Writer::default();
        bucket.varint(1, 1);
        objects.push(object(
            identifier,
            HEADER_BUCKET_MESSAGE_TYPE,
            bucket.finish(),
            Vec::new(),
        ));
    }
    let lists = [
        (Some(ids.string_list), LIST_STRING),
        (Some(ids.style_list), LIST_STYLE),
        (Some(ids.formula_list), LIST_FORMULA),
        (Some(ids.format_list), LIST_FORMAT),
        (ids.control_cell_spec_list, LIST_CONTROL_CELL_SPEC),
    ];
    for (identifier, list_type) in lists {
        if let Some(identifier) = identifier {
            objects.push(object(
                identifier,
                DATA_LIST_MESSAGE_TYPE,
                data_list(list_type),
                Vec::new(),
            ));
        }
    }
    objects.push(object(
        ids.uid_map,
        UID_MAP_MESSAGE_TYPE,
        super::empty_storage(
            super::Role::UidMap,
            &[],
            shape.rows,
            shape.columns,
            uid_seed,
        )?,
        Vec::new(),
    ));
    objects.push(object(
        ids.stroke_sidecar,
        STROKE_SIDECAR_MESSAGE_TYPE,
        super::empty_storage(super::Role::Stroke, &[], shape.rows, shape.columns, 0)?,
        Vec::new(),
    ));
    Ok(objects)
}

/// A table style preset naming `network`.
#[must_use]
pub fn preset(identifier: u64, index: u32, network: u64) -> Object {
    let mut preset = Writer::default();
    preset.varint(1, u64::from(index)).reference(3, network);
    object(
        identifier,
        PRESET_MESSAGE_TYPE,
        preset.finish(),
        vec![network],
    )
}

/// A table style network naming `styles`, with `preset_id` the styles'
/// preset index.
#[must_use]
pub fn style_network(identifier: u64, styles: &Styles) -> Object {
    let mut network = Writer::default();
    network
        .reference(1, styles.body_text)
        .reference(2, styles.header_row_text)
        .reference(3, styles.header_column_text)
        .reference(4, styles.footer_row_text)
        .reference(5, styles.body_cell)
        .reference(6, styles.header_row_cell)
        .reference(7, styles.header_column_cell)
        .reference(8, styles.footer_row_cell)
        .reference(9, styles.table);
    if let Some(name) = styles.table_name {
        network.reference(10, name);
    }
    if let Some(shape) = styles.table_name_shape {
        network.reference(11, shape);
    }
    network.varint(12, u64::from(styles.preset_index));
    // The host declares the text styles, the name styles, the cell styles
    // and the table style, each once.
    let mut references = Vec::with_capacity(11);
    let candidates = [
        Some(styles.body_text),
        Some(styles.header_row_text),
        Some(styles.header_column_text),
        Some(styles.footer_row_text),
        styles.table_name,
        styles.table_name_shape,
        Some(styles.body_cell),
        Some(styles.header_row_cell),
        Some(styles.header_column_cell),
        Some(styles.footer_row_cell),
        Some(styles.table),
    ];
    for identifier in candidates.into_iter().flatten() {
        if !references.contains(&identifier) {
            references.push(identifier);
        }
    }
    object(
        identifier,
        STYLE_NETWORK_MESSAGE_TYPE,
        network.finish(),
        references,
    )
}

/// The style network a table style preset names.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a preset without one network.
pub fn preset_network(preset: &[u8]) -> Result<u64> {
    super::single_reference(preset, 3)?
        .filter(|identifier| *identifier != 0)
        .ok_or(WriteError::Invalid)
}

/// The stylesheet context the preset `preset` gives a new table: the styles
/// its style network names, and the preset's own index (else
/// `theme_index`, its position among the theme's presets), as the host
/// editor's `theme_table_style` reads them.
///
/// # Errors
///
/// [`WriteError::Invalid`] for a malformed preset or network, or one that
/// leaves a required style unnamed.
pub fn preset_styles(
    preset: u64,
    theme_index: u32,
    preset_payload: &[u8],
    network: &[u8],
) -> Result<Styles> {
    let required = |number: u32| -> Result<u64> {
        super::single_reference(network, number)?
            .filter(|identifier| *identifier != 0)
            .ok_or(WriteError::Invalid)
    };
    let optional = |number: u32| -> Result<Option<u64>> {
        Ok(super::single_reference(network, number)?.filter(|identifier| *identifier != 0))
    };
    let index = match litchi_iwa_common::wire::repeated_varint_values(preset_payload, 1)
        .map_err(|_error| WriteError::Invalid)?
        .last()
    {
        // An `int32`: a negative index is refused, as the host refuses it.
        Some(value) => u32::try_from(*value).map_err(|_error| WriteError::Invalid)?,
        None => theme_index,
    };
    Ok(Styles {
        preset,
        preset_index: index,
        table: required(9)?,
        body_text: required(1)?,
        header_row_text: required(2)?,
        header_column_text: required(3)?,
        footer_row_text: required(4)?,
        body_cell: required(5)?,
        header_row_cell: required(6)?,
        header_column_cell: required(7)?,
        footer_row_cell: required(8)?,
        table_name: optional(10)?,
        table_name_shape: optional(11)?,
    })
}

/// The identifiers of a fresh default table stylesheet ([`default_styles`]).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StyleIds {
    /// The stylesheet holding the styles.
    pub stylesheet: u64,
    /// The list style.
    pub list: u64,
    /// The paragraph style (body, table name and every text region).
    pub paragraph: u64,
    /// The character style.
    pub character: u64,
    /// The shape style (the table name's).
    pub shape: u64,
    /// The media style.
    pub media: u64,
    /// The drop-cap style.
    pub drop_cap: u64,
    /// The sheet style.
    pub sheet: u64,
    /// The table style.
    pub table: u64,
    /// The cell style.
    pub cell: u64,
}

impl StyleIds {
    /// The styles, in the host's order (the stylesheet's style list).
    #[must_use]
    pub const fn styles(&self) -> [u64; 9] {
        [
            self.list,
            self.paragraph,
            self.character,
            self.shape,
            self.media,
            self.drop_cap,
            self.sheet,
            self.table,
            self.cell,
        ]
    }

    /// The [`Styles`] context these styles give a table with `preset`.
    #[must_use]
    pub const fn context(&self, preset: u64) -> Styles {
        Styles::uniform(preset, self.table, self.paragraph, self.cell, self.shape)
    }
}

fn style(name: &str, identifier: &str, stylesheet: u64) -> Vec<u8> {
    let mut style = Writer::default();
    style
        .bytes(1, name.as_bytes())
        .bytes(2, identifier.as_bytes())
        .reference(5, stylesheet);
    style.finish()
}

/// A style message: `super` and the override count at `count_field`, then
/// the `extra` fields.
fn style_message(name: &str, identifier: &str, stylesheet: u64, count_field: u32) -> Writer {
    let mut message = Writer::default();
    message
        .bytes(1, &style(name, identifier, stylesheet))
        .varint(count_field, 0);
    message
}

/// The host Numbers builder's default table styles in a stylesheet of their
/// own, as the host Pages builder imports them for `body_table`: the
/// stylesheet (its styles and the body paragraph style's identifier entry)
/// first, then each style, all declaring the stylesheet.
#[must_use]
pub fn default_styles(ids: &StyleIds) -> Vec<Object> {
    let sheet = ids.stylesheet;
    let mut stylesheet = Writer::default();
    for identifier in ids.styles() {
        stylesheet.reference(1, identifier);
    }
    stylesheet
        .message(2, |entry| {
            entry
                .bytes(1, b"litchi-paragraph-body")
                .reference(2, ids.paragraph);
        })
        .bool(4, false)
        .bool(6, true);
    let mut objects = vec![object(
        sheet,
        STYLESHEET_MESSAGE_TYPE,
        stylesheet.finish(),
        ids.styles().to_vec(),
    )];
    let simple = |name: &str, identifier: &str, field: u32| {
        style_message(name, identifier, sheet, field).finish()
    };
    let mut shape = Writer::default();
    shape
        .bytes(1, &simple("Shape", "litchi-shape-default", 10))
        .varint(10, 0);
    let mut table = style_message("Table", "litchi-table-default", sheet, 10);
    table.message(11, |properties| {
        properties.bool(21, true).bool(22, false);
    });
    let mut cell = style_message("Cell", "litchi-cell-default", sheet, 10);
    cell.message(11, |properties| {
        properties.bool(3, false);
    });
    for (identifier, message_type, payload) in [
        (
            ids.list,
            LIST_STYLE_MESSAGE_TYPE,
            simple("None", "litchi-list-none", 10),
        ),
        (
            ids.character,
            CHARACTER_STYLE_MESSAGE_TYPE,
            simple("Default", "litchi-character-default", 10),
        ),
        (ids.shape, SHAPE_STYLE_MESSAGE_TYPE, shape.finish()),
        (
            ids.media,
            MEDIA_STYLE_MESSAGE_TYPE,
            simple("Media", "litchi-media-default", 10),
        ),
        (
            ids.drop_cap,
            DROP_CAP_STYLE_MESSAGE_TYPE,
            simple("Drop Cap", "litchi-dropcap", 10),
        ),
        (
            ids.sheet,
            SHEET_STYLE_MESSAGE_TYPE,
            simple("Sheet", "litchi-sheet-default", 2),
        ),
        (ids.table, TABLE_STYLE_MESSAGE_TYPE, table.finish()),
        (ids.cell, CELL_STYLE_MESSAGE_TYPE, cell.finish()),
        (
            ids.paragraph,
            PARAGRAPH_STYLE_MESSAGE_TYPE,
            simple("Body", "litchi-paragraph-body", 10),
        ),
    ] {
        objects.push(object(identifier, message_type, payload, vec![sheet]));
    }
    objects
}

/// A calculation engine whose only formula owner is `owner` (internal
/// identifier `internal`, UID `uid`), saved under `locale`, as the host
/// Numbers builder writes it.
#[must_use]
pub fn engine(identifier: u64, owner: u64, internal: u32, uid: (u64, u64), locale: &str) -> Object {
    let mut engine = Writer::default();
    engine
        .message(2, |tracker| {
            tracker
                .message(3, |map| {
                    map.message(1, |entry| {
                        entry.varint(1, u64::from(internal)).message(2, |cfuuid| {
                            cfuuid
                                .varint(2, uid.0 & 0xffff_ffff)
                                .varint(3, uid.0 >> 32)
                                .varint(4, uid.1 & 0xffff_ffff)
                                .varint(5, uid.1 >> 32);
                        });
                    });
                })
                .varint(5, 0)
                .reference(6, owner);
        })
        .bytes(16, locale.as_bytes());
    object(
        identifier,
        CALCULATION_ENGINE_MESSAGE_TYPE,
        engine.finish(),
        vec![owner],
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn shape(rows: u32, columns: u32) -> Shape<'static> {
        Shape {
            name: "Table 1",
            rows,
            columns,
            row_height: None,
            column_width: None,
        }
    }

    #[test]
    fn identifiers_follow_the_host_order() {
        let builder = Ids::consecutive(9, Profile::Builder);
        assert_eq!(builder.all(), (9..20).collect::<Vec<_>>());
        assert_eq!(builder.count(), 11);
        let bootstrap = Ids::consecutive(100, Profile::Bootstrap);
        assert_eq!(bootstrap.control_cell_spec_list, Some(109));
        assert_eq!((bootstrap.uid_map, bootstrap.stroke_sidecar), (110, 111));
        assert_eq!(bootstrap.all().len(), bootstrap.count());
    }

    #[test]
    fn graphs_are_deterministic_and_checked() {
        let ids = Ids::consecutive(9, Profile::Builder);
        let styles = Styles::uniform(120, 118, 112, 119, 114);
        let placement = Placement {
            parent: 8,
            frame: DEFAULT_FRAME,
        };
        let uuid = super::super::TableUuid::derive(1).text;
        let first = table(
            &ids,
            placement,
            &styles,
            &shape(10, 5),
            &uuid,
            10,
            Profile::Builder,
        )
        .unwrap_or_else(|error| panic!("{error:?}"));
        let again = table(
            &ids,
            placement,
            &styles,
            &shape(10, 5),
            &uuid,
            10,
            Profile::Builder,
        )
        .unwrap_or_else(|error| panic!("{error:?}"));
        assert_eq!(first, again);
        assert_eq!(first.len(), 11);
        assert_eq!(first[0].references, vec![8, 9 + 1]);
        assert_eq!(
            first[1].references,
            [118, 112, 114, 119, 120]
                .into_iter()
                .chain(11..20)
                .collect::<Vec<_>>()
        );
        assert!(
            table(
                &ids,
                placement,
                &styles,
                &shape(0, 5),
                &uuid,
                10,
                Profile::Builder
            )
            .is_err()
        );
        assert!(
            table(
                &ids,
                placement,
                &styles,
                &shape(3, 3),
                &uuid,
                10,
                Profile::Bootstrap
            )
            .is_err()
        );
        let bootstrap = Ids::consecutive(9, Profile::Bootstrap);
        let objects = table(
            &bootstrap,
            placement,
            &styles,
            &shape(3, 3),
            &uuid,
            10,
            Profile::Bootstrap,
        )
        .unwrap_or_else(|error| panic!("{error:?}"));
        let mut sorted = objects[1].references.clone();
        sorted.sort_unstable();
        assert_eq!(objects[1].references, sorted);
        assert_eq!(objects.len(), 12);
    }

    #[test]
    fn default_styles_declare_their_stylesheet() {
        let ids = StyleIds {
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
        let objects = default_styles(&ids);
        assert_eq!(objects.len(), 10);
        assert_eq!(objects[0].references, (11..20).collect::<Vec<_>>());
        assert!(objects[1..].iter().all(|object| object.references == [40]));
        let context = ids.context(20);
        assert_eq!(context.objects(), vec![20, 18, 12, 14, 19]);
    }
}
