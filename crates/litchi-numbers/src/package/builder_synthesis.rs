//! One-pass builder options over the blank template (record 0370i).
//!
//! A fresh workbook with options is the template with a handful of fields
//! rewritten, so the builder writes them straight into the template's two
//! changed components — `Index/Document.iwa` and
//! `Index/CalculationEngine.iwa` — and opens the result once, instead of
//! publishing and reopening one structural edit per option. The fields are
//! exactly those the transactions it replaces write:
//!
//! - names (the `names` transaction): the sheet's `name` and the table
//!   model's `table_name`;
//! - dimensions (`table::resize`): the model's size, the tiles' removed
//!   rows and extent (`maxColumn`, `maxRow`, as the host builder writes
//!   them), the header buckets' headers inside the new size, a fresh
//!   row/column UID map (the template table stores no cell), the stroke
//!   sidecar's size and the table owner's table and body extents;
//! - language and locale (`document_locale`): the document's language,
//!   current and creation locale and the calculation engine's locale.
//!
//! The package metadata is untouched: the output is one first save, with
//! the template's save tokens.

use std::collections::HashMap;
use std::ops::Range;
use std::sync::Arc;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::varint::{decode_varint_from_bytes, encode_varint_into};
use litchi_iwa_common::wire::{append_length_delimited_field, parse_wire_view, patch_varint_field};
use litchi_iwa_core::SnappyStream;

use super::document_locale;
use super::structure::Fault;
use super::table_topology::wire::single_bytes;
use super::table_topology::{self as topology, TILE_MESSAGE_TYPE};
use super::{Package, Profile};
use litchi_numbers_wire::table_create::{Shape, fresh};

const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
const ENGINE_COMPONENT: &str = "Index/CalculationEngine.iwa";
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHEET_MESSAGE_TYPE: u32 = 2;
const MODEL_MESSAGE_TYPE: u32 = 6_001;
const UID_MAP_MESSAGE_TYPE: u32 = 6_267;
const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
/// The template's sheet, which parents its table.
const TEMPLATE_SHEET: u64 = 8;
/// The template table's identifiers (the host builder's).
const TEMPLATE_TABLE: fresh::Ids = fresh::Ids {
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
/// The template table's styles: preset, table, paragraph, cell and
/// table-name shape style.
const TEMPLATE_STYLES: fresh::Styles = fresh::Styles::uniform(20, 18, 12, 19, 14);
const ENGINE_MESSAGE_TYPE: u32 = 4_000;
const OWNER_MESSAGE_TYPE: u32 = 4_008;
const SHEET_NAME_FIELD: u32 = 1;
const MODEL_NAME_FIELD: u32 = 8;
const MODEL_ROWS_FIELD: u32 = 6;
const MODEL_COLUMNS_FIELD: u32 = 7;
/// The most row and column identities one table may hold.
const MAX_TABLE_UIDS: u64 = 1_100_000;

/// The options one build writes; `None` keeps the template's value.
#[derive(Debug, Clone, Copy, Default)]
pub(crate) struct Options<'a> {
    pub(crate) sheet_name: Option<&'a str>,
    pub(crate) table_name: Option<&'a str>,
    pub(crate) dimensions: Option<(u32, u32)>,
    pub(crate) language: Option<&'a str>,
    pub(crate) locale: Option<&'a str>,
}

/// Why a synthesis failed.
#[derive(Debug)]
pub(crate) enum Failure {
    /// The requested size leaves no body row or column, or needs too many
    /// identities.
    InvalidSize,
    /// The template is not the shape this synthesis rewrites.
    Template,
    /// The rewritten package could not be assembled or opened.
    Package(crate::Error),
}

impl From<Fault> for Failure {
    fn from(_fault: Fault) -> Self {
        Self::Template
    }
}

impl From<topology::Error> for Failure {
    fn from(_error: topology::Error) -> Self {
        Self::Template
    }
}

impl From<litchi_numbers_wire::table_topology::Error> for Failure {
    fn from(_error: litchi_numbers_wire::table_topology::Error) -> Self {
        Self::Template
    }
}

fn core(error: litchi_iwa_core::Error) -> Failure {
    Failure::Package(crate::Error::Archive(litchi_iwa_archive::Error::Iwa(error)))
}

const fn wire(_error: &litchi_iwa_common::Error) -> Failure {
    Failure::Template
}

/// One object of a component stream: its header and its payloads.
struct Object {
    identifier: u64,
    /// The whole object: length prefix, header and payloads.
    span: Range<usize>,
    header: Range<usize>,
    messages: Vec<(u32, Range<usize>)>,
}

/// A decompressed IWA component, scanned object by object, with the
/// payloads a synthesis replaces. Only the object headers' identifiers and
/// message types and lengths are read; the final open validates the
/// result in full.
struct Stream {
    bytes: Vec<u8>,
    objects: Vec<Object>,
    replaced: HashMap<(usize, usize), Vec<u8>>,
}

/// The varint value of a field's payload.
fn leading_varint(payload: &[u8]) -> Result<u64, Failure> {
    decode_varint_from_bytes(payload)
        .map(|(value, _)| value)
        .map_err(|_error| Failure::Template)
}

impl Stream {
    fn scan(bytes: Vec<u8>) -> Result<Self, Failure> {
        let mut objects = Vec::new();
        let mut position = 0;
        while position < bytes.len() {
            let (length, prefix) =
                decode_varint_from_bytes(&bytes[position..]).map_err(|_error| Failure::Template)?;
            let start = position + prefix;
            let end = start
                .checked_add(usize::try_from(length).map_err(|_error| Failure::Template)?)
                .filter(|end| *end <= bytes.len())
                .ok_or(Failure::Template)?;
            let header = &bytes[start..end];
            let mut identifier = None;
            let mut infos = Vec::new();
            for field in parse_wire_view(header, litchi_iwa_common::WireLimits::DEFAULT)
                .map_err(|error| wire(&error))?
                .fields()
            {
                match field.number() {
                    1 => identifier = Some(leading_varint(field.payload())?),
                    2 => infos.push(field.payload()),
                    _ => {},
                }
            }
            let identifier = identifier.ok_or(Failure::Template)?;
            let mut messages = Vec::new();
            let mut cursor = end;
            for info in infos {
                let (mut type_, mut length) = (None, None);
                for field in parse_wire_view(info, litchi_iwa_common::WireLimits::DEFAULT)
                    .map_err(|error| wire(&error))?
                    .fields()
                {
                    match field.number() {
                        1 => type_ = Some(leading_varint(field.payload())?),
                        3 => length = Some(leading_varint(field.payload())?),
                        _ => {},
                    }
                }
                let type_ = u32::try_from(type_.ok_or(Failure::Template)?)
                    .map_err(|_error| Failure::Template)?;
                let length = usize::try_from(length.ok_or(Failure::Template)?)
                    .map_err(|_error| Failure::Template)?;
                let next = cursor
                    .checked_add(length)
                    .filter(|next| *next <= bytes.len())
                    .ok_or(Failure::Template)?;
                messages.push((type_, cursor..next));
                cursor = next;
            }
            objects.push(Object {
                identifier,
                span: position..cursor,
                header: start..end,
                messages,
            });
            position = cursor;
        }
        Ok(Self {
            bytes,
            objects,
            replaced: HashMap::new(),
        })
    }

    /// Every message of `message_type`.
    fn of_type(&self, message_type: u32) -> Vec<(usize, usize)> {
        let mut output = Vec::new();
        for (object_index, object) in self.objects.iter().enumerate() {
            for (message_index, (type_, _)) in object.messages.iter().enumerate() {
                if *type_ == message_type {
                    output.push((object_index, message_index));
                }
            }
        }
        output
    }

    /// The message of `message_type` of object `identifier`.
    fn locate(&self, identifier: u64, message_type: u32) -> Result<(usize, usize), Failure> {
        for (object_index, object) in self.objects.iter().enumerate() {
            if object.identifier == identifier {
                for (message_index, (type_, _)) in object.messages.iter().enumerate() {
                    if *type_ == message_type {
                        return Ok((object_index, message_index));
                    }
                }
            }
        }
        Err(Failure::Template)
    }

    fn message(&self, location: (usize, usize)) -> Result<&[u8], Failure> {
        if let Some(data) = self.replaced.get(&location) {
            return Ok(data);
        }
        let (_, range) = self
            .objects
            .get(location.0)
            .and_then(|object| object.messages.get(location.1))
            .ok_or(Failure::Template)?;
        Ok(&self.bytes[range.clone()])
    }

    fn set(&mut self, location: (usize, usize), data: Vec<u8>) {
        self.replaced.insert(location, data);
    }

    /// The component bytes with every replacement and its message length.
    fn finish(self) -> Result<Vec<u8>, Failure> {
        if self.replaced.is_empty() {
            return Ok(self.bytes);
        }
        let mut output = Vec::with_capacity(self.bytes.len() + 64);
        for (object_index, object) in self.objects.iter().enumerate() {
            let changed = (0..object.messages.len())
                .any(|message| self.replaced.contains_key(&(object_index, message)));
            if !changed {
                output.extend_from_slice(&self.bytes[object.span.clone()]);
                continue;
            }
            let header = &self.bytes[object.header.clone()];
            let mut rewritten = Vec::with_capacity(header.len() + 8);
            let mut info_index = 0;
            let view = parse_wire_view(header, litchi_iwa_common::WireLimits::DEFAULT)
                .map_err(|error| wire(&error))?;
            for field in view.fields() {
                if field.number() != 2 {
                    rewritten.extend_from_slice(field.raw());
                    continue;
                }
                match self.replaced.get(&(object_index, info_index)) {
                    Some(data) => {
                        let length =
                            u64::try_from(data.len()).map_err(|_error| Failure::Template)?;
                        let info = patch_varint_field(field.payload(), 3, true, Some(length))
                            .map_err(|error| wire(&error))?;
                        append_length_delimited_field(
                            &mut rewritten,
                            2,
                            &info,
                            litchi_iwa_common::WireLimits::DEFAULT,
                        )
                        .map_err(|error| wire(&error))?;
                    },
                    None => rewritten.extend_from_slice(field.raw()),
                }
                info_index += 1;
            }
            encode_varint_into(
                &mut output,
                u64::try_from(rewritten.len()).map_err(|_error| Failure::Template)?,
            );
            output.extend_from_slice(&rewritten);
            for (message_index, (_, range)) in object.messages.iter().enumerate() {
                match self.replaced.get(&(object_index, message_index)) {
                    Some(data) => output.extend_from_slice(data),
                    None => output.extend_from_slice(&self.bytes[range.clone()]),
                }
            }
        }
        Ok(output)
    }
}

fn set_string(data: &[u8], number: u32, value: &str) -> Result<Vec<u8>, Failure> {
    let present = single_bytes(data, number)?.is_some();
    litchi_iwa_common::wire::patch_length_delimited_field(
        data,
        number,
        present,
        Some(value.as_bytes()),
    )
    .map_err(|error| wire(&error))
}

/// The model's size and section counts and its data store, read in one
/// pass (`table_topology::geometry` rereads the model per field).
fn geometry(model: &[u8]) -> Result<topology::Geometry, Failure> {
    let mut values = [None; 5];
    for field in parse_wire_view(model, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|error| wire(&error))?
        .fields()
    {
        let slot = match field.number() {
            MODEL_ROWS_FIELD => 0,
            MODEL_COLUMNS_FIELD => 1,
            9 => 2,
            10 => 3,
            11 => 4,
            _ => continue,
        };
        values[slot] = Some(
            u32::try_from(leading_varint(field.payload())?).map_err(|_error| Failure::Template)?,
        );
    }
    let geometry = topology::Geometry {
        rows: values[0].ok_or(Failure::Template)?,
        columns: values[1].ok_or(Failure::Template)?,
        header_rows: values[2].unwrap_or(0),
        header_columns: values[3].unwrap_or(0),
        footer_rows: values[4].unwrap_or(0),
    };
    if geometry.header_rows + geometry.footer_rows > geometry.rows
        || geometry.header_columns > geometry.columns
    {
        return Err(Failure::Template);
    }
    Ok(geometry)
}

/// The document component with every option written.
#[allow(
    clippy::too_many_lines,
    reason = "one pass over the template's document component"
)]
fn document(
    stream: &mut Stream,
    options: &Options<'_>,
) -> Result<Option<(topology::Geometry, u32, u32)>, Failure> {
    if options.language.is_some() || options.locale.is_some() {
        let [location] = stream.of_type(DOCUMENT_MESSAGE_TYPE)[..] else {
            return Err(Failure::Template);
        };
        let data =
            document_locale::document(stream.message(location)?, options.language, options.locale)?;
        stream.set(location, data);
    }
    if let Some(name) = options.sheet_name {
        let [location] = stream.of_type(SHEET_MESSAGE_TYPE)[..] else {
            return Err(Failure::Template);
        };
        let data = set_string(stream.message(location)?, SHEET_NAME_FIELD, name)?;
        stream.set(location, data);
    }
    let [model_location] = stream.of_type(MODEL_MESSAGE_TYPE)[..] else {
        return Err(Failure::Template);
    };
    let resize = match options.dimensions {
        Some((rows, columns)) => {
            if rows == 0 || columns == 0 || u64::from(rows) + u64::from(columns) > MAX_TABLE_UIDS {
                return Err(Failure::InvalidSize);
            }
            let before = geometry(stream.message(model_location)?)?;
            if rows <= before.header_rows + before.footer_rows || columns <= before.header_columns {
                return Err(Failure::InvalidSize);
            }
            ((before.rows, before.columns) != (rows, columns)).then_some((before, rows, columns))
        },
        None => None,
    };
    if options.table_name.is_none() && resize.is_none() {
        return Ok(None);
    }
    // The table is the shared from-scratch graph (record 0420) with the
    // template's identifiers, styles and identity: the template is the
    // host builder's table, which the graph reproduces byte for byte.
    let model = stream.message(model_location)?.to_vec();
    let before = geometry(&model)?;
    let (rows, columns) = resize.map_or((before.rows, before.columns), |(_, rows, columns)| {
        (rows, columns)
    });
    let uuid = std::str::from_utf8(single_bytes(&model, 1)?.ok_or(Failure::Template)?)
        .map_err(|_error| Failure::Template)?
        .to_owned();
    let template_name =
        std::str::from_utf8(single_bytes(&model, MODEL_NAME_FIELD)?.ok_or(Failure::Template)?)
            .map_err(|_error| Failure::Template)?
            .to_owned();
    let shape = Shape {
        name: options.table_name.unwrap_or(&template_name),
        rows,
        columns,
        row_height: None,
        column_width: None,
    };
    let objects = fresh::table(
        &TEMPLATE_TABLE,
        fresh::Placement {
            parent: TEMPLATE_SHEET,
            frame: fresh::DEFAULT_FRAME,
        },
        &TEMPLATE_STYLES,
        &shape,
        &uuid,
        TEMPLATE_TABLE.model,
        fresh::Profile::Builder,
    )
    .map_err(|_error| Failure::Template)?;
    for object in objects {
        // The drawable, the buckets and the lists do not depend on the
        // size or the name; the UID map is the template's own until the
        // table is resized.
        let changes = match object.message_type {
            MODEL_MESSAGE_TYPE => true,
            TILE_MESSAGE_TYPE | UID_MAP_MESSAGE_TYPE | STROKE_SIDECAR_MESSAGE_TYPE => {
                resize.is_some()
            },
            _ => false,
        };
        if changes {
            let location = stream.locate(object.identifier, object.message_type)?;
            stream.set(location, object.payload);
        }
    }
    Ok(resize)
}

/// The engine component with the locale and the owner's extents written.
fn engine(
    stream: &mut Stream,
    locale: Option<&str>,
    resized: Option<(topology::Geometry, u32, u32)>,
) -> Result<(), Failure> {
    if let Some(locale) = locale {
        for location in stream.of_type(ENGINE_MESSAGE_TYPE) {
            let data = document_locale::engine(stream.message(location)?, locale)?;
            stream.set(location, data);
        }
    }
    if let Some((before, rows, columns)) = resized {
        let extents = |rows: u32, columns: u32| {
            topology::calc::Extents::of(
                rows,
                columns,
                before.header_rows,
                before.header_columns,
                before.footer_rows,
            )
        };
        for location in stream.of_type(OWNER_MESSAGE_TYPE) {
            if let Some(rewritten) = topology::calc::resize_extents(
                stream.message(location)?,
                extents(before.rows, before.columns),
                extents(rows, columns),
            )? {
                stream.set(location, rewritten);
            }
        }
    }
    Ok(())
}

/// The template `source` with `identity` (when not the template's own)
/// and `options` written, opened once.
pub(crate) fn synthesize(
    source: &[u8],
    identity: Option<crate::create::Identity>,
    options: &Options<'_>,
) -> Result<Package, Failure> {
    let profile = Profile::default();
    let limits = profile.archive();
    let catalog =
        Catalog::from_bytes(source, limits).map_err(|error| Failure::Package(error.into()))?;
    let snappy = limits
        .snappy_limits()
        .map_err(|error| Failure::Package(error.into()))?;
    let scan = |name: &str| -> Result<Stream, Failure> {
        let entry = catalog
            .iter()
            .find(|entry| entry.name() == name && !entry.is_opaque())
            .ok_or(Failure::Template)?;
        let stream = SnappyStream::decompress(entry.data(), snappy).map_err(core)?;
        Stream::scan(stream.into_bytes())
    };
    let compress = |bytes: Vec<u8>| -> Result<Vec<u8>, Failure> {
        SnappyStream::compress(&bytes, snappy).map_err(core)
    };
    let mut members: Vec<(&str, Vec<u8>)> = Vec::new();
    if let Some(identity) = identity {
        members
            .extend(crate::create::identity_members(&catalog, identity).map_err(Failure::Package)?);
    }
    let mut document_stream = scan(DOCUMENT_COMPONENT)?;
    let resized = document(&mut document_stream, options)?;
    members.push((DOCUMENT_COMPONENT, compress(document_stream.finish()?)?));
    if options.locale.is_some() || resized.is_some() {
        let mut engine_stream = scan(ENGINE_COMPONENT)?;
        engine(&mut engine_stream, options.locale, resized)?;
        members.push((ENGINE_COMPONENT, compress(engine_stream.finish()?)?));
    }
    let edits: Vec<EntryEdit<'_>> = members
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect();
    let bytes = catalog
        .reassemble_with_deletions_to_bytes(&edits, &[], limits)
        .map_err(|error| Failure::Package(error.into()))?;
    Package::from_shared_profile(Arc::from(bytes), profile).map_err(Failure::Package)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The template's table is the shared from-scratch graph (record 0420):
    /// every object it builds, apart from the host builder's random UID
    /// map, is the template's own payload byte for byte, and so are the
    /// preset and style network.
    #[test]
    fn the_template_table_is_the_shared_fresh_graph() {
        let profile = Profile::default();
        let limits = profile.archive();
        let catalog = Catalog::from_bytes(crate::create::BLANK, limits)
            .unwrap_or_else(|error| panic!("catalog: {error:?}"));
        let snappy = limits
            .snappy_limits()
            .unwrap_or_else(|error| panic!("limits: {error:?}"));
        let entry = catalog
            .iter()
            .find(|entry| entry.name() == DOCUMENT_COMPONENT)
            .unwrap_or_else(|| panic!("document"));
        let bytes = SnappyStream::decompress(entry.data(), snappy)
            .unwrap_or_else(|error| panic!("snappy: {error:?}"))
            .into_bytes();
        let stream = Stream::scan(bytes).unwrap_or_else(|_| panic!("scan"));
        let model = stream
            .message(stream.locate(10, MODEL_MESSAGE_TYPE).unwrap_or((0, 0)))
            .unwrap_or_else(|_| panic!("model"))
            .to_vec();
        let uuid = std::str::from_utf8(single_bytes(&model, 1).ok().flatten().unwrap_or_default())
            .unwrap_or_default()
            .to_owned();
        let objects = fresh::table(
            &TEMPLATE_TABLE,
            fresh::Placement {
                parent: TEMPLATE_SHEET,
                frame: fresh::DEFAULT_FRAME,
            },
            &TEMPLATE_STYLES,
            &Shape {
                name: "Table 1",
                rows: 10,
                columns: 5,
                row_height: None,
                column_width: None,
            },
            &uuid,
            10,
            fresh::Profile::Builder,
        )
        .unwrap_or_else(|error| panic!("fresh: {error:?}"));
        let mut compared = 0;
        for object in objects.into_iter().chain([
            fresh::preset(20, 0, 32),
            fresh::style_network(32, &TEMPLATE_STYLES),
        ]) {
            if object.message_type == fresh::UID_MAP_MESSAGE_TYPE {
                continue;
            }
            let location = stream
                .locate(object.identifier, object.message_type)
                .unwrap_or_else(|_| panic!("object {}", object.identifier));
            assert_eq!(
                stream.message(location).ok(),
                Some(object.payload.as_slice()),
                "object {}",
                object.identifier
            );
            compared += 1;
        }
        assert_eq!(compared, 12);
    }
}
