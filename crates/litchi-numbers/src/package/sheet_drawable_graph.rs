//! Private proof of the drawables a rooted Numbers sheet owns.
//!
//! The sheet-drawable families share this module: it resolves a sheet from
//! the document root, reads its back-to-front drawable list
//! (`TN.SheetArchive.drawable_infos`), classifies each drawable by its
//! native message and selects one by a [`DrawableSelector`]. Native
//! identifiers never leave the package adapter; the families keep them in
//! their private targets.

use std::collections::HashSet;

use litchi_core::Position;
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{WireFieldView, WireView};
use litchi_iwa_core::ArchiveObject;
use litchi_iwa_protos::shape_properties_codec as codec;

use super::Package;
use super::structure::{
    self, DOCUMENT_COMPONENT, DOCUMENT_SHEETS_FIELD, Fault, FaultResult, references,
};
use crate::selector::SheetSelector;
use crate::sheet::drawable::{Drawable, DrawableSelector, Geometry, GeometryFlags, Kind};
use crate::sheet::drawable::{Point, Properties, Size};

/// `TSWP.ShapeInfoArchive`.
pub(super) const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
/// `TSD.ImageArchive`.
pub(super) const IMAGE_MESSAGE_TYPE: u32 = 3_005;
/// `TSD.MovieArchive`.
pub(super) const MOVIE_MESSAGE_TYPE: u32 = 3_007;
/// `TSCH.ChartDrawableArchive`.
pub(super) const CHART_MESSAGE_TYPE: u32 = 5_021;
/// `TSWP.ShapeInfoArchive.is_text_box`.
const SHAPE_INFO_IS_TEXT_BOX_FIELD: u32 = 6;
/// `TSD.MovieArchive.audioOnly`.
const MOVIE_AUDIO_ONLY_FIELD: u32 = 9;
/// `TSD.MovieArchive.is_live_video`.
const MOVIE_LIVE_VIDEO_FIELD: u32 = 30;
/// The `super` field every wrapped drawable message starts with.
const WRAPPED_DRAWABLE_FIELD: u32 = 1;
/// `TSD.DrawableArchive.parent`.
const DRAWABLE_PARENT_FIELD: u32 = 2;

/// The rooted sheet a family edits.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct SheetTarget {
    /// The sheet's zero-based workbook position.
    pub(super) position: usize,
    /// The sheet object.
    pub(super) identifier: u64,
    /// The sheet message inside the sheet object.
    pub(super) message_index: usize,
    /// The field path of the drawable list in that message.
    pub(super) path: &'static [u32],
}

/// One drawable of the sheet list.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Entry {
    /// Its object.
    pub(super) identifier: u64,
    /// What it is.
    pub(super) kind: Kind,
    /// The message carrying its drawable (`None` for a table or another
    /// kind without a classified message).
    pub(super) message: Option<(usize, u32)>,
    /// Its zero-based list position.
    pub(super) position: usize,
    /// Its zero-based position among drawables of its kind.
    pub(super) kind_position: usize,
}

/// A sheet's resolved drawable list.
#[derive(Debug, Clone)]
pub(super) struct SheetDrawables {
    pub(super) sheet: SheetTarget,
    pub(super) entries: Vec<Entry>,
}

impl SheetDrawables {
    /// The entry `selector` names, if any.
    pub(super) fn select(&self, selector: DrawableSelector) -> Option<&Entry> {
        match selector {
            DrawableSelector::Position(position) => self.entries.get(position.get()),
            DrawableSelector::Of { kind, position } => self
                .entries
                .iter()
                .find(|entry| entry.kind == kind && entry.kind_position == position.get()),
        }
    }

    /// The drawable identifiers in list order.
    pub(super) fn identifiers(&self) -> Vec<u64> {
        self.entries.iter().map(|entry| entry.identifier).collect()
    }
}

/// The current object `identifier` of `package` (the first physical copy;
/// exact aliases are coalesced by the package index).
pub(super) fn object(package: &Package, identifier: u64) -> FaultResult<&ArchiveObject> {
    let resolved = package
        .state
        .index
        .resolve_ref_id(&package.state.components, identifier)
        .map_err(|_error| Fault::Invalid)?
        .ok_or(Fault::Invalid)?;
    package
        .state
        .components
        .catalog()
        .get_index(resolved.component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .ok_or(Fault::Invalid)
}

/// The zero-based position of the sheet `selector` names, or `None`.
pub(super) fn sheet_position(
    package: &Package,
    selector: SheetSelector<'_>,
) -> FaultResult<Option<usize>> {
    match selector {
        SheetSelector::Name(name) => {
            let mut matches = package
                .workbook_ref()
                .sheets()
                .iter()
                .filter(|sheet| sheet.name() == name);
            let Some(sheet) = matches.next() else {
                return Ok(None);
            };
            if matches.next().is_some() {
                return Err(Fault::Invalid);
            }
            Ok(Some(sheet.index()))
        },
        SheetSelector::Position(position) => Ok(package
            .workbook_ref()
            .sheets()
            .get(position.get())
            .map(crate::Sheet::index)),
    }
}

/// Resolve the rooted sheet at workbook `position`.
pub(super) fn sheet_target(package: &Package, position: usize) -> FaultResult<SheetTarget> {
    let document = package
        .state
        .components
        .get_archive(DOCUMENT_COMPONENT)
        .and_then(|archive| archive.object(1))
        .ok_or(Fault::Invalid)?;
    let document_index = structure::document_message(document)?;
    let sheets = references::identifiers(
        &document.messages[document_index].data,
        &[DOCUMENT_SHEETS_FIELD],
    )?;
    let identifier = *sheets.get(position).ok_or(Fault::Invalid)?;
    let sheet = object(package, identifier)?;
    let (message_index, path) = structure::sheet_message(sheet)?;
    Ok(SheetTarget {
        position,
        identifier,
        message_index,
        path,
    })
}

/// Resolve and classify the drawables of the sheet at `position`.
pub(super) fn sheet_drawables(package: &Package, position: usize) -> FaultResult<SheetDrawables> {
    let sheet = sheet_target(package, position)?;
    let sheet_object = object(package, sheet.identifier)?;
    let identifiers =
        references::identifiers(&sheet_object.messages[sheet.message_index].data, sheet.path)?;
    let mut seen = HashSet::new();
    let mut entries = Vec::new();
    structure::reserve(&mut entries, identifiers.len())?;
    let mut counts = [0_usize; 8];
    for (list_position, identifier) in identifiers.into_iter().enumerate() {
        if identifier == sheet.identifier || identifier == 1 || !seen.insert(identifier) {
            return Err(Fault::Invalid);
        }
        let drawable = object(package, identifier)?;
        let (kind, message) = classify(drawable)?;
        let slot = kind_slot(kind);
        let kind_position = counts[slot];
        counts[slot] += 1;
        entries.push(Entry {
            identifier,
            kind,
            message,
            position: list_position,
            kind_position,
        });
    }
    Ok(SheetDrawables { sheet, entries })
}

const fn kind_slot(kind: Kind) -> usize {
    match kind {
        Kind::TextBox => 0,
        Kind::Shape => 1,
        Kind::Image => 2,
        Kind::Movie => 3,
        Kind::Audio => 4,
        Kind::Chart => 5,
        Kind::Table => 6,
        Kind::Other => 7,
    }
}

/// The kind of one sheet drawable and the message that carries it.
fn classify(object: &ArchiveObject) -> FaultResult<(Kind, Option<(usize, u32)>)> {
    if structure::table_info_message(object).is_ok() {
        return Ok((Kind::Table, None));
    }
    for message_type in [
        SHAPE_INFO_MESSAGE_TYPE,
        IMAGE_MESSAGE_TYPE,
        MOVIE_MESSAGE_TYPE,
        CHART_MESSAGE_TYPE,
    ] {
        let mut found = None;
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ == message_type && found.replace(index).is_some() {
                return Err(Fault::Invalid);
            }
        }
        let Some(index) = found else {
            continue;
        };
        let payload = object.messages[index].data.as_slice();
        let kind = match message_type {
            SHAPE_INFO_MESSAGE_TYPE => match bool_field(payload, SHAPE_INFO_IS_TEXT_BOX_FIELD)? {
                Some(true) => Kind::TextBox,
                Some(false) => Kind::Shape,
                None => Kind::Other,
            },
            IMAGE_MESSAGE_TYPE => Kind::Image,
            MOVIE_MESSAGE_TYPE => {
                if bool_field(payload, MOVIE_LIVE_VIDEO_FIELD)? == Some(true) {
                    Kind::Other
                } else if bool_field(payload, MOVIE_AUDIO_ONLY_FIELD)? == Some(true) {
                    Kind::Audio
                } else {
                    Kind::Movie
                }
            },
            _ => Kind::Chart,
        };
        return Ok((kind, Some((index, message_type))));
    }
    Ok((Kind::Other, None))
}

/// The last occurrence of boolean field `number` (protobuf merge rule).
fn bool_field(payload: &[u8], number: u32) -> FaultResult<Option<bool>> {
    let view = WireView::parse(payload, WireLimits::DEFAULT)
        .map_err(|error| structure::map_wire(&error))?;
    let mut value = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 0 {
            return Err(Fault::Invalid);
        }
        let (raw, _) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
            .map_err(|_error| Fault::Invalid)?;
        value = Some(raw != 0);
    }
    Ok(value)
}

/// The single field `number` of `view`.
pub(super) fn unique_field<'a>(
    view: &WireView<'a>,
    number: u32,
    wire_type: u8,
) -> FaultResult<Option<WireFieldView<'a>>> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != wire_type {
            return Err(Fault::Invalid);
        }
        if found.replace(field).is_some() {
            return Err(Fault::Invalid);
        }
    }
    Ok(found)
}

/// The `TSD.DrawableArchive` inside a drawable message: field 1 of an image,
/// movie or chart, field 1 of field 1 of a `TSWP.ShapeInfoArchive`.
pub(super) fn drawable_payload(message_type: u32, payload: &[u8]) -> FaultResult<&[u8]> {
    let view = WireView::parse(payload, WireLimits::DEFAULT)
        .map_err(|error| structure::map_wire(&error))?;
    let outer = unique_field(&view, WRAPPED_DRAWABLE_FIELD, 2)?
        .ok_or(Fault::Invalid)?
        .payload();
    if message_type != SHAPE_INFO_MESSAGE_TYPE {
        return Ok(outer);
    }
    let shape =
        WireView::parse(outer, WireLimits::DEFAULT).map_err(|error| structure::map_wire(&error))?;
    Ok(unique_field(&shape, WRAPPED_DRAWABLE_FIELD, 2)?
        .ok_or(Fault::Invalid)?
        .payload())
}

/// The parent a `TSD.DrawableArchive` names, if any.
pub(super) fn drawable_parent(drawable: &[u8]) -> FaultResult<Option<u64>> {
    let view = WireView::parse(drawable, WireLimits::DEFAULT)
        .map_err(|error| structure::map_wire(&error))?;
    unique_field(&view, DRAWABLE_PARENT_FIELD, 2)?
        .map(|field| references::reference_identifier(field.payload()))
        .transpose()
}

/// Codec options sized to one drawable payload.
pub(super) fn codec_options(source: &[u8]) -> codec::DecodeOptions {
    codec::DecodeOptions::for_source(source)
}

/// The frame and properties a `TSD.DrawableArchive` stores.
pub(super) fn frame_and_properties(drawable: &[u8]) -> FaultResult<(Option<Geometry>, Properties)> {
    let decoded = codec::decode_drawable_archive(drawable, codec_options(drawable))
        .map_err(|_error| Fault::Invalid)?;
    let geometry = decoded.geometry().map(|geometry| Geometry {
        position: geometry.position().map(|point| Point {
            x: point.x(),
            y: point.y(),
        }),
        size: geometry.size().map(|size| Size {
            width: size.width(),
            height: size.height(),
        }),
        flags: geometry.flags().map(GeometryFlags::from_bits),
        angle: geometry.angle(),
    });
    if geometry.is_some_and(|geometry| !geometry.is_valid()) {
        return Err(Fault::Invalid);
    }
    let stored = decoded.properties();
    let properties = Properties {
        hyperlink: stored.hyperlink_url().map(str::to_owned),
        locked: stored.locked(),
        aspect_ratio_locked: stored.aspect_ratio_locked(),
        accessibility_description: stored.accessibility_description().map(str::to_owned),
    };
    Ok((geometry, properties))
}

/// The public list of the sheet's drawables.
pub(super) fn list(package: &Package, drawables: &SheetDrawables) -> FaultResult<Vec<Drawable>> {
    let mut output = Vec::new();
    structure::reserve(&mut output, drawables.entries.len())?;
    for entry in &drawables.entries {
        let geometry = match entry.message {
            Some((index, message_type)) => {
                let drawable_object = object(package, entry.identifier)?;
                let payload = drawable_object.messages[index].data.as_slice();
                frame_and_properties(drawable_payload(message_type, payload)?)?.0
            },
            None => None,
        };
        output.push(Drawable {
            position: Position::new(entry.position),
            kind: entry.kind,
            kind_position: Position::new(entry.kind_position),
            geometry,
        });
    }
    Ok(output)
}

/// One in-place change of an existing object: `edit` rewrites the object's
/// messages (and their header references) under the archive limits.
pub(super) type ObjectEdit<'e> =
    &'e mut dyn FnMut(&mut ArchiveObject, litchi_iwa_core::Limits) -> FaultResult<()>;

/// Publish a candidate in which only the given objects changed in place,
/// keeping every other member's exact bytes, the root previews and the
/// package metadata, and reopen it through the strict ingress.
///
/// Each edited component is cloned from the source's parsed archive (never
/// re-parsed); an object stored in more than one component is refused.
pub(super) fn publish_in_place(
    source: &Package,
    edits: &mut [(u64, ObjectEdit<'_>)],
) -> FaultResult<structure::Published> {
    let catalog = structure::exact_source(source)?;
    let archive_limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(|error| structure::map_archive(&error))?;
    let components = source.state.components.catalog();
    let mut archives: Vec<(String, litchi_iwa_core::Archive)> = Vec::new();
    for (identifier, edit) in edits.iter_mut() {
        let mut owner = None;
        for component in components.iter() {
            if component.archive().object(*identifier).is_some()
                && owner.replace(component).is_some()
            {
                return Err(Fault::Unsupported);
            }
        }
        let owner = owner.ok_or(Fault::Invalid)?;
        let slot = match archives.iter().position(|(name, _)| name == owner.name()) {
            Some(slot) => slot,
            None => {
                archives.push((owner.name().to_owned(), owner.archive().clone()));
                archives.len() - 1
            },
        };
        let object = archives[slot]
            .1
            .object_mut(*identifier)
            .ok_or(Fault::Invalid)?;
        edit(object, archive_limits)?;
    }
    let snappy = catalog
        .limits()
        .snappy_limits()
        .map_err(|error| structure::map_archive(&error))?;
    let mut encoded = Vec::new();
    structure::reserve(&mut encoded, archives.len())?;
    for (name, archive) in &archives {
        let bytes = archive
            .to_bytes(archive_limits)
            .map_err(|error| structure::map_core(&error))?;
        let compressed = litchi_iwa_core::SnappyStream::compress(&bytes, snappy)
            .map_err(|error| structure::map_core(&error))?;
        encoded.push((name.as_str(), compressed));
    }
    let entries: Vec<litchi_iwa_archive::package::EntryEdit<'_>> = encoded
        .iter()
        .map(|(name, data)| litchi_iwa_archive::package::EntryEdit::new(name, data))
        .collect();
    let output = catalog
        .package()
        .reassemble_with_changes_to_bytes(&[], &entries, &[], catalog.limits())
        .map_err(|error| structure::map_archive(&error))?;
    let package = source
        .reopen_candidate(output.into())
        .map_err(|_error| Fault::Verification)?;
    Ok(structure::Published {
        package,
        touched_components: encoded.len(),
        deleted_previews: 0,
    })
}

/// What a selector names for a text family.
pub(super) enum TextTarget {
    /// The storage the text box or shape owns, and its storage message.
    Storage {
        /// The storage object.
        storage: u64,
    },
    /// The drawable owns no text.
    NoText(Kind),
    /// No drawable matches.
    NotFound,
}

/// The storage the text box or shape `selector` names owns.
pub(super) fn text_storage(
    package: &Package,
    sheet: usize,
    selector: DrawableSelector,
) -> FaultResult<TextTarget> {
    let drawables = sheet_drawables(package, sheet)?;
    let Some(entry) = drawables.select(selector) else {
        return Ok(TextTarget::NotFound);
    };
    if !entry.kind.has_text() {
        return Ok(TextTarget::NoText(entry.kind));
    }
    let (index, message_type) = entry.message.ok_or(Fault::Invalid)?;
    if message_type != SHAPE_INFO_MESSAGE_TYPE {
        return Err(Fault::Invalid);
    }
    let shape = object(package, entry.identifier)?;
    let records = references::records(&shape.messages[index].data, &[4])?;
    let [record] = records.as_slice() else {
        return Err(Fault::Invalid);
    };
    let storage = references::reference_identifier(record)?;
    structure::unique_message(object(package, storage)?, &[2_001])?;
    Ok(TextTarget::Storage { storage })
}
