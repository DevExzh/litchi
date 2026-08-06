//! Native body-footnote CRUD for Pages documents.

use std::collections::HashSet;

use prost::Message;

use super::text_box_create::body_text_storage;
use super::{
    DOCUMENT_OBJECT_ID, PagesEditor, STORAGE_MESSAGE_TYPES, find_object_archive,
    package_references_object,
};
use crate::archive::{ArchiveObject, RawMessage};
use crate::package_metadata::{
    add_component_object_uuids, component_identifier_for_object_uuid, next_object_identifier,
    release_package_identifier_suffix, remove_component_external_references_to_object,
    remove_component_object_uuids, set_package_last_object_identifier,
};
use crate::protobuf::{tsp, tswp};
use crate::text::IWorkTextEditor;
use crate::text::editor::storage_object_references;
use crate::wire::{
    patch_length_delimited_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use crate::{Error, IWorkPackage, Result};
use litchi_pages::footnote::body::{Footnote, Position, Selector};

const FOOTNOTE_REFERENCE_MESSAGE_TYPE: u32 = 2_008;
const TEXTUAL_ATTACHMENT_MESSAGE_TYPE: u32 = 2_004;
const FOOTNOTE_TABLE_FIELD: u32 = 16;
const TABLE_ENTRIES_FIELD: u32 = 1;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const FOOTNOTE_ANCHOR: char = '\u{000e}';
const FOOTNOTE_ANCHOR_TEXT: &str = "\u{000e}";
const FOOTNOTE_ANCHOR_UNIT: u16 = 0x000e;
const FOOTNOTE_MARK: char = '\u{fffc}';
const FOOTNOTE_CONTENT_PREFIX: &str = "\u{fffc} ";

/// Native Pages footnote data plus the private objects it owns.
#[derive(Debug, Clone)]
pub(super) struct BodyFootnoteGraph {
    pub(super) footnote: Footnote,
    reference_id: u64,
    storage_id: u64,
    marker_id: u64,
}

#[derive(Debug)]
struct FootnoteTableEntry {
    index: u32,
    reference_id: u64,
    raw: Vec<u8>,
}

#[derive(Debug, Clone, Copy)]
struct FootnoteObjectIds {
    reference: u64,
    storage: u64,
    marker: u64,
}

impl FootnoteObjectIds {
    fn allocate(first: u64) -> Result<Self> {
        let identifier = |offset| {
            first.checked_add(offset).ok_or_else(|| {
                Error::ParseError("Pages footnote object identifier overflow".to_owned())
            })
        };
        Ok(Self {
            reference: identifier(0)?,
            storage: identifier(1)?,
            marker: identifier(2)?,
        })
    }

    const fn last(self) -> u64 {
        self.marker
    }
}

impl PagesEditor {
    /// Read every native footnote attached to the main Pages body.
    pub fn body_footnotes(&self) -> Result<Vec<Footnote>> {
        Ok(
            body_footnote_graphs(self.package(), self.body_storage_id.get())?
                .into_iter()
                .map(|graph| graph.footnote)
                .collect(),
        )
    }

    /// Insert a native Pages footnote at a UTF-16 body position.
    ///
    /// The inserted body character is Pages' private U+000E footnote anchor;
    /// use [`Self::body_footnotes`] instead of treating that character as text.
    pub fn insert_body_footnote(
        &mut self,
        position: Position,
        text: impl AsRef<str>,
    ) -> Result<Footnote> {
        let text = text.as_ref();
        validate_footnote_text(text)?;
        let position_u32 = position.utf16_index();
        let position_index = usize::try_from(position_u32).map_err(|_| {
            Error::ParseError("Pages footnote position exceeds the platform index range".to_owned())
        })?;
        body_footnote_graphs(self.package(), self.body_storage_id.get())?;

        let mut text_editor = IWorkTextEditor::from_package(self.package().clone());
        text_editor.replace_text(
            self.body_storage_id,
            position_index..position_index,
            FOOTNOTE_ANCHOR_TEXT,
        )?;
        let mut staged = text_editor.into_package();
        let ids = FootnoteObjectIds::allocate(next_object_identifier(&staged)?)?;
        let body = storage_at(&staged, self.body_storage_id.get(), "Pages body")?.1;
        let archive_name = find_object_archive(&staged, self.body_storage_id.get())?;
        let objects = new_footnote_objects(ids, text, &body)?;

        insert_footnote_reference(
            &mut staged,
            &archive_name,
            self.body_storage_id.get(),
            position_u32,
            ids.reference,
        )?;
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &[ids.storage])?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = body_footnote_by_selector(&verified, Selector::At(position))?.footnote;
        if created.position.utf16_index() != position_u32
            || created.text.as_ref() != text
            || created.custom_mark.is_some()
        {
            return Err(Error::InvalidFormat(
                "Pages footnote insertion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Replace the user-visible text of one native body footnote.
    pub fn set_body_footnote_text(
        &mut self,
        selector: Selector,
        text: impl AsRef<str>,
    ) -> Result<Footnote> {
        let text = text.as_ref();
        validate_footnote_text(text)?;
        let current = body_footnote_by_selector(self, selector)?;
        if current.footnote.text.as_ref() == text {
            return Ok(current.footnote);
        }
        let storage = storage_at(self.package(), current.storage_id, "Pages footnote")?.1;
        let content = storage.text.concat();
        let prefix_units = FOOTNOTE_CONTENT_PREFIX.encode_utf16().count();
        let content_units = content.encode_utf16().count();
        if content_units < prefix_units {
            return Err(Error::InvalidFormat(format!(
                "Pages footnote storage {} is shorter than its native marker",
                current.storage_id
            )));
        }

        let mut text_editor = IWorkTextEditor::from_package(self.package().clone());
        text_editor.replace_text(
            crate::text::native_storage_id(current.storage_id)?,
            prefix_units..content_units,
            text,
        )?;
        let verified = Self::from_bytes(&text_editor.into_package().to_bytes()?)?;
        let updated = body_footnote_by_selector(&verified, selector)?.footnote;
        if updated.position != current.footnote.position
            || updated.text.as_ref() != text
            || updated.custom_mark != current.footnote.custom_mark
        {
            return Err(Error::InvalidFormat(
                "Pages footnote text update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(updated)
    }

    /// Delete one native body footnote, its body anchor, and its owned objects.
    pub fn remove_body_footnote(&mut self, selector: Selector) -> Result<Footnote> {
        let removed = body_footnote_by_selector(self, selector)?.footnote;
        let start = usize::try_from(removed.position.utf16_index()).map_err(|_| {
            Error::ParseError("Pages footnote position exceeds the platform index range".to_owned())
        })?;
        let end = start
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Pages footnote anchor range overflow".to_owned()))?;
        self.replace_body_text(start..end, "")?;
        if self
            .body_footnotes()?
            .iter()
            .any(|footnote| footnote.position == removed.position)
        {
            return Err(Error::InvalidFormat(
                "Pages footnote deletion failed validation".to_owned(),
            ));
        }
        Ok(removed)
    }
}

pub(super) fn body_footnote_graphs(
    package: &IWorkPackage,
    body_storage_id: u64,
) -> Result<Vec<BodyFootnoteGraph>> {
    let (_, body, body_data) = storage_at_with_data(package, body_storage_id, "Pages body")?;
    let entries = footnote_table_entries(body_storage_id, &body_data, &body)?;
    let mut seen = HashSet::new();
    let mut footnotes = Vec::with_capacity(entries.len());
    for entry in entries {
        if !seen.insert(entry.reference_id) {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {body_storage_id} references footnote object {} more than once",
                entry.reference_id
            )));
        }
        footnotes.push(decode_footnote_graph(
            package,
            entry.index,
            entry.reference_id,
        )?);
    }
    Ok(footnotes)
}

/// Reclaim footnote graphs whose anchors were removed by an ordinary body edit.
pub(super) fn cleanup_removed_body_footnotes(
    package: &mut IWorkPackage,
    body_storage_id: u64,
    before: &[BodyFootnoteGraph],
) -> Result<()> {
    if before.is_empty() {
        return Ok(());
    }
    let remaining = body_footnote_graphs(package, body_storage_id)?
        .into_iter()
        .map(|graph| graph.reference_id)
        .collect::<HashSet<_>>();
    let removed = before
        .iter()
        .filter(|graph| !remaining.contains(&graph.reference_id))
        .collect::<Vec<_>>();
    if removed.is_empty() {
        return Ok(());
    }

    let mut staged = package.clone();
    let mut identifiers = Vec::with_capacity(removed.len() * 3);
    for graph in removed {
        identifiers.extend(remove_unreferenced_footnote_graph(&mut staged, graph)?);
    }
    release_package_identifier_suffix(&mut staged, &identifiers)?;
    IWorkPackage::from_bytes(&staged.to_bytes()?)?;
    *package = staged;
    Ok(())
}

fn body_footnote_by_selector(
    editor: &PagesEditor,
    selector: Selector,
) -> Result<BodyFootnoteGraph> {
    let footnotes = body_footnote_graphs(editor.package(), editor.body_storage_id.get())?;
    match selector {
        Selector::Index(index) => footnotes.into_iter().nth(index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages body has no footnote at source index {index}"
            ))
        }),
        Selector::At(position) => {
            let mut matches = footnotes
                .into_iter()
                .filter(|graph| graph.footnote.position == position);
            let Some(graph) = matches.next() else {
                return Err(Error::InvalidFormat(format!(
                    "Pages body has no footnote at UTF-16 position {}",
                    position.utf16_index()
                )));
            };
            if matches.next().is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Pages body has more than one footnote at UTF-16 position {}",
                    position.utf16_index()
                )));
            }
            Ok(graph)
        },
    }
}

fn decode_footnote_graph(
    package: &IWorkPackage,
    position: u32,
    reference_id: u64,
) -> Result<BodyFootnoteGraph> {
    let reference_archive = find_object_archive(package, reference_id)?;
    let reference_archive_data = package.archive(&reference_archive)?;
    let reference_object = reference_archive_data.object(reference_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages footnote object {reference_id} is missing"))
    })?;
    let reference = tswp::FootnoteReferenceAttachmentArchive::decode(object_message_data(
        reference_object,
        FOOTNOTE_REFERENCE_MESSAGE_TYPE,
        "Pages footnote reference",
    )?)?;
    if reference
        .super_
        .as_ref()
        .and_then(|value| value.kind)
        .is_some_and(|kind| {
            kind != tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32
        })
    {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote object {reference_id} has the wrong attachment kind"
        )));
    }
    let storage_id = reference
        .contained_storage
        .as_ref()
        .map(|value| value.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages footnote object {reference_id} has no contained storage"
            ))
        })?;
    let (_, storage, _) = storage_at_with_data(package, storage_id, "Pages footnote")?;
    if storage.kind != Some(tswp::storage_archive::KindType::Footnote as i32) {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote storage {storage_id} is not a native footnote storage"
        )));
    }
    let content = storage.text.concat();
    let text = content
        .strip_prefix(FOOTNOTE_CONTENT_PREFIX)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages footnote storage {storage_id} lacks its native marker prefix"
            ))
        })?;
    let marker_id = footnote_marker_id(storage_id, &storage)?;
    validate_footnote_marker(package, marker_id)?;

    let position = Position::from_utf16_index(usize::try_from(position).map_err(|_| {
        Error::ParseError("Pages footnote position exceeds the platform index range".to_owned())
    })?)
    .map_err(|error| Error::ParseError(format!("invalid Pages footnote position: {error}")))?;
    let footnote =
        Footnote::with_custom_mark(position, text, reference.custom_mark_string.map(Into::into))
            .map_err(|error| Error::ParseError(format!("invalid Pages footnote value: {error}")))?;

    Ok(BodyFootnoteGraph {
        footnote,
        reference_id,
        storage_id,
        marker_id,
    })
}

fn validate_footnote_marker(package: &IWorkPackage, marker_id: u64) -> Result<()> {
    let archive_name = find_object_archive(package, marker_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(marker_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages footnote marker object {marker_id} is missing"
        ))
    })?;
    let marker = tswp::TextualAttachmentArchive::decode(object_message_data(
        object,
        TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
        "Pages footnote marker",
    )?)?;
    if marker.kind != Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32) {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote marker object {marker_id} has the wrong attachment kind"
        )));
    }
    Ok(())
}

fn footnote_marker_id(storage_id: u64, storage: &tswp::StorageArchive) -> Result<u64> {
    let entries = storage
        .table_attachment
        .as_ref()
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages footnote storage {storage_id} has no marker attachment table"
            ))
        })?
        .entries
        .iter()
        .filter(|entry| entry.character_index == 0)
        .collect::<Vec<_>>();
    let [entry] = entries.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote storage {storage_id} must have exactly one marker attachment at index zero"
        )));
    };
    entry
        .object
        .as_ref()
        .map(|value| value.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages footnote storage {storage_id} has an invalid marker attachment"
            ))
        })
}

fn new_footnote_objects(
    ids: FootnoteObjectIds,
    text: &str,
    body: &tswp::StorageArchive,
) -> Result<[ArchiveObject; 3]> {
    let mut content = String::with_capacity(FOOTNOTE_CONTENT_PREFIX.len() + text.len());
    content.push_str(FOOTNOTE_CONTENT_PREFIX);
    content.push_str(text);
    let mut storage = body_text_storage(&content, body);
    storage.kind = Some(tswp::storage_archive::KindType::Footnote as i32);
    storage.table_attachment = Some(tswp::ObjectAttributeTable {
        entries: vec![tswp::object_attribute_table::ObjectAttribute {
            character_index: 0,
            object: Some(reference(ids.marker)),
        }],
    });
    let marker = tswp::TextualAttachmentArchive {
        string_equivalent: None,
        kind: Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32),
    };
    let attachment = tswp::FootnoteReferenceAttachmentArchive {
        super_: None,
        contained_storage: Some(reference(ids.storage)),
        custom_mark_string: None,
    };
    let storage_references = storage_object_references(&storage);
    Ok([
        pages_object(
            ids.reference,
            FOOTNOTE_REFERENCE_MESSAGE_TYPE,
            attachment,
            &[ids.storage],
        )?,
        pages_object(
            ids.storage,
            STORAGE_MESSAGE_TYPES[0],
            storage,
            &storage_references,
        )?,
        pages_object(ids.marker, TEXTUAL_ATTACHMENT_MESSAGE_TYPE, marker, &[])?,
    ])
}

fn pages_object(
    identifier: u64,
    message_type: u32,
    message: impl Message,
    references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: message.encode_to_vec(),
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references = references.to_vec();
    Ok(object)
}

fn insert_footnote_reference(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    position: u32,
    reference_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages body storage {storage_id} is missing"))
        })?;
        let message_index = unique_storage_message_index(object, storage_id)?;
        let original = &object.messages[message_index];
        let storage = tswp::StorageArchive::decode(original.data.as_slice())?;
        let tables = repeated_length_delimited_payloads(original.data.as_slice(), FOOTNOTE_TABLE_FIELD)?;
        if tables.len() > 1 {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} contains {} footnote tables",
                tables.len()
            )));
        }
        let mut entries = footnote_table_entries(storage_id, original.data.as_slice(), &storage)?;
        if entries.iter().any(|entry| entry.index == position) {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} already has a footnote at UTF-16 index {position}"
            )));
        }
        require_text_boundary(storage_id, position, &storage.text)?;
        if utf16_unit_at(&storage.text, position) != Some(FOOTNOTE_ANCHOR_UNIT) {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} has no U+000E footnote anchor at UTF-16 index {position}"
            )));
        }
        let new_entry = tswp::object_attribute_table::ObjectAttribute {
            character_index: position,
            object: Some(reference(reference_id)),
        };
        entries.push(FootnoteTableEntry {
            index: position,
            reference_id,
            raw: new_entry.encode_to_vec(),
        });
        entries.sort_by_key(|entry| entry.index);
        let encoded_entries = entries
            .into_iter()
            .map(|entry| entry.raw)
            .collect::<Vec<_>>();
        let table = match tables.first() {
            Some(table) => rewrite_repeated_length_delimited_fields(
                table,
                TABLE_ENTRIES_FIELD,
                &encoded_entries,
            )?,
            None => tswp::ObjectAttributeTable {
                entries: encoded_entries
                    .iter()
                    .map(|entry| tswp::object_attribute_table::ObjectAttribute::decode(entry.as_slice()))
                    .collect::<std::result::Result<Vec<_>, _>>()?,
            }
            .encode_to_vec(),
        };
        let data = patch_length_delimited_field(
            original.data.as_slice(),
            FOOTNOTE_TABLE_FIELD,
            !tables.is_empty(),
            Some(&table),
        )?;
        let verified = tswp::StorageArchive::decode(data.as_slice())?;
        if footnote_table_entries(storage_id, &data, &verified)?
            .iter()
            .all(|entry| entry.reference_id != reference_id)
        {
            return Err(Error::InvalidFormat(
                "Pages body footnote table patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        let references = &mut object.archive_info.message_infos[message_index].object_references;
        if references.contains(&reference_id) {
            return Err(Error::InvalidFormat(format!(
                "Pages body metadata already references footnote object {reference_id}"
            )));
        }
        references.push(reference_id);
        Ok(())
    })
}

fn footnote_table_entries(
    storage_id: u64,
    data: &[u8],
    storage: &tswp::StorageArchive,
) -> Result<Vec<FootnoteTableEntry>> {
    let tables = repeated_length_delimited_payloads(data, FOOTNOTE_TABLE_FIELD)?;
    let [table] = tables.as_slice() else {
        return match tables.len() {
            0 if storage.table_footnote.is_none() => Ok(Vec::new()),
            0 => Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} footnote table wire state is inconsistent"
            ))),
            count => Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} contains {count} footnote tables"
            ))),
        };
    };
    if storage.table_footnote.is_none() {
        return Err(Error::InvalidFormat(format!(
            "Pages body storage {storage_id} footnote table wire state is inconsistent"
        )));
    }
    let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?
        .into_iter()
        .map(|raw| {
            let entry = tswp::object_attribute_table::ObjectAttribute::decode(raw)?;
            Ok(FootnoteTableEntry {
                index: entry.character_index,
                reference_id: entry.object.map_or(0, |value| value.identifier),
                raw: raw.to_vec(),
            })
        })
        .collect::<Result<Vec<_>>>()?;
    validate_footnote_table_entries(storage_id, &entries, &storage.text)?;
    Ok(entries)
}

fn validate_footnote_table_entries(
    storage_id: u64,
    entries: &[FootnoteTableEntry],
    text: &[String],
) -> Result<()> {
    let text_length = text_utf16_len(text)?;
    let mut previous = None;
    for entry in entries {
        if entry.reference_id == 0 {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} has a zero footnote object identifier"
            )));
        }
        if previous.is_some_and(|index| index >= entry.index) {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} footnote positions are not strictly increasing"
            )));
        }
        require_text_boundary(storage_id, entry.index, text)?;
        if entry.index >= text_length
            || utf16_unit_at(text, entry.index) != Some(FOOTNOTE_ANCHOR_UNIT)
        {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} footnote {} is not anchored to U+000E at UTF-16 index {}",
                entry.reference_id, entry.index
            )));
        }
        previous = Some(entry.index);
    }
    Ok(())
}

fn remove_unreferenced_footnote_graph(
    package: &mut IWorkPackage,
    graph: &BodyFootnoteGraph,
) -> Result<Vec<u64>> {
    let reference_id = graph.reference_id;
    remove_unreferenced_footnote_object(
        package,
        reference_id,
        &[FOOTNOTE_REFERENCE_MESSAGE_TYPE],
        "reference attachment",
    )?;
    remove_unreferenced_footnote_object(
        package,
        graph.storage_id,
        STORAGE_MESSAGE_TYPES,
        "storage",
    )?;
    remove_unreferenced_footnote_object(
        package,
        graph.marker_id,
        &[TEXTUAL_ATTACHMENT_MESSAGE_TYPE],
        "marker attachment",
    )?;
    Ok(vec![reference_id, graph.storage_id, graph.marker_id])
}

fn remove_unreferenced_footnote_object(
    package: &mut IWorkPackage,
    identifier: u64,
    message_types: &[u32],
    label: &str,
) -> Result<()> {
    if package_references_object(package, identifier)? {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote {label} object {identifier} remains referenced after body-anchor deletion"
        )));
    }
    remove_component_external_references_to_object(package, DOCUMENT_OBJECT_ID, identifier)?;
    if let Some(component) = component_identifier_for_object_uuid(package, identifier)? {
        remove_component_object_uuids(package, component, &[identifier])?;
    }
    let archive_name = find_object_archive(package, identifier)?;
    package.update_archive(&archive_name, |archive| {
        let object = archive.remove_object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages footnote {label} object {identifier} is missing"
            ))
        })?;
        object_message_data_of_types(&object, message_types, &format!("Pages footnote {label}"))?;
        Ok(())
    })
}

fn storage_at(
    package: &IWorkPackage,
    storage_id: u64,
    label: &str,
) -> Result<(String, tswp::StorageArchive)> {
    let (archive_name, storage, _) = storage_at_with_data(package, storage_id, label)?;
    Ok((archive_name, storage))
}

fn storage_at_with_data(
    package: &IWorkPackage,
    storage_id: u64,
    label: &str,
) -> Result<(String, tswp::StorageArchive, Vec<u8>)> {
    let archive_name = find_object_archive(package, storage_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive
        .object(storage_id)
        .ok_or_else(|| Error::InvalidFormat(format!("{label} storage {storage_id} is missing")))?;
    let message_index = unique_storage_message_index(object, storage_id)?;
    let data = object.messages[message_index].data.clone();
    Ok((
        archive_name,
        tswp::StorageArchive::decode(data.as_slice())?,
        data,
    ))
}

fn unique_storage_message_index(object: &ArchiveObject, storage_id: u64) -> Result<usize> {
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| STORAGE_MESSAGE_TYPES.contains(&message.type_))
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let [index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages text storage {storage_id} must have exactly one writable payload"
        )));
    };
    Ok(*index)
}

fn object_message_data<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    label: &str,
) -> Result<&'a [u8]> {
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == message_type)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{label} must contain exactly one message type {message_type}"
        )));
    };
    Ok(message.data.as_slice())
}

fn object_message_data_of_types<'a>(
    object: &'a ArchiveObject,
    message_types: &[u32],
    label: &str,
) -> Result<&'a [u8]> {
    let messages = object
        .messages
        .iter()
        .filter(|message| message_types.contains(&message.type_))
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{label} must contain exactly one supported message payload"
        )));
    };
    Ok(message.data.as_slice())
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_type: None,
        deprecated_is_external: None,
    }
}

fn validate_footnote_text(text: &str) -> Result<()> {
    if text.contains(FOOTNOTE_ANCHOR) || text.contains(FOOTNOTE_MARK) {
        return Err(Error::ParseError(
            "Pages footnote text cannot contain native footnote-anchor or attachment markers"
                .to_owned(),
        ));
    }
    Ok(())
}

fn require_text_boundary(storage_id: u64, position: u32, text: &[String]) -> Result<()> {
    let mut current = 0u32;
    if position == current {
        return Ok(());
    }
    for fragment in text {
        for character in fragment.chars() {
            current = current
                .checked_add(character.len_utf16() as u32)
                .ok_or_else(|| {
                    Error::InvalidFormat("Pages text UTF-16 length overflow".to_owned())
                })?;
            if current == position {
                return Ok(());
            }
            if current > position {
                break;
            }
        }
    }
    Err(Error::InvalidFormat(format!(
        "UTF-16 index {position} is not a scalar boundary in Pages storage {storage_id}"
    )))
}

fn text_utf16_len(text: &[String]) -> Result<u32> {
    text.iter().try_fold(0u32, |total, fragment| {
        fragment.chars().try_fold(total, |total, character| {
            total
                .checked_add(character.len_utf16() as u32)
                .ok_or_else(|| Error::InvalidFormat("Pages text UTF-16 length overflow".to_owned()))
        })
    })
}

fn utf16_unit_at(text: &[String], requested: u32) -> Option<u16> {
    let mut index = 0u32;
    for fragment in text {
        for unit in fragment.encode_utf16() {
            if index == requested {
                return Some(unit);
            }
            index = index.checked_add(1)?;
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_pages::footnote::body::Footnote;

    #[test]
    fn body_footnote_crud_round_trips_and_restores_a_source_document() {
        let mut editor = PagesEditor::create_with_text("A😀B").unwrap();
        let baseline = editor.to_bytes().unwrap();
        let note = editor
            .insert_body_footnote(Position::from_utf16_index(3).unwrap(), "Initial note")
            .unwrap();
        assert_eq!(editor.body_text().unwrap(), "A😀\u{e}B");
        assert_eq!(note.position, Position::from_utf16_index(3).unwrap());
        assert_eq!(note.text.as_ref(), "Initial note");
        assert_eq!(note.custom_mark, None);
        assert_eq!(editor.body_footnotes().unwrap(), vec![note.clone()]);

        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.body_footnotes().unwrap(), vec![note.clone()]);

        let updated = editor
            .set_body_footnote_text(Selector::Index(0), "Updated note")
            .unwrap();
        assert_eq!(updated.text.as_ref(), "Updated note");
        assert_eq!(updated.position, note.position);

        let removed = editor
            .remove_body_footnote(Selector::At(note.position))
            .unwrap();
        assert_eq!(removed, updated);
        assert_eq!(editor.body_text().unwrap(), "A😀B");
        assert!(editor.body_footnotes().unwrap().is_empty());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn ordinary_body_replacement_reclaims_deleted_footnote_graphs() {
        let mut editor = PagesEditor::create_with_text("AB").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(1).unwrap(), "First")
            .unwrap();
        let first_reference_id =
            body_footnote_graphs(editor.package(), editor.body_storage_id.get()).unwrap()[0]
                .reference_id;
        let second = editor
            .insert_body_footnote(Position::from_utf16_index(3).unwrap(), "Second")
            .unwrap();
        assert_eq!(editor.body_text().unwrap(), "A\u{e}B\u{e}");

        editor.replace_body_text(1..2, "").unwrap();
        assert_eq!(editor.body_text().unwrap(), "AB\u{e}");
        assert_eq!(
            editor.body_footnotes().unwrap(),
            vec![Footnote {
                position: Position::from_utf16_index(2).unwrap(),
                text: "Second".into(),
                custom_mark: None,
            }]
        );
        assert!(find_object_archive(editor.package(), first_reference_id).is_err());
        assert_eq!(second.position, Position::from_utf16_index(3).unwrap());
    }

    #[test]
    fn footnote_text_rejects_native_structural_markers_transactionally() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(
            editor
                .insert_body_footnote(Position::ZERO, "Invalid\u{e}")
                .is_err()
        );
        assert!(
            editor
                .insert_body_footnote(Position::ZERO, "Invalid\u{fffc}")
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn native_footnote_reference_without_a_super_payload_is_supported() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let footnote = editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let reference_id = body_footnote_graphs(editor.package(), editor.body_storage_id.get())
            .unwrap()[0]
            .reference_id;
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, reference_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(reference_id).unwrap();
                let message = &object.messages[0];
                let mut attachment =
                    tswp::FootnoteReferenceAttachmentArchive::decode(message.data.as_slice())?;
                attachment.super_ = None;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: FOOTNOTE_REFERENCE_MESSAGE_TYPE,
                        data: attachment.encode_to_vec(),
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let parsed = PagesEditor::from_package(package).unwrap();
        assert_eq!(parsed.body_footnotes().unwrap(), vec![footnote]);
    }
}
