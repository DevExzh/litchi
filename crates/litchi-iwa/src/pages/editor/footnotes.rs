//! Native body-footnote CRUD for Pages documents.

use std::collections::HashSet;
use std::hash::Hash;

use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::pages_body_codec;
use litchi_iwa_protos::pages_footnote_codec;
use litchi_iwa_protos::pages_footnote_marker_codec;
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
#[cfg(test)]
const FOOTNOTE_SUPER_FIELD: u32 = 1;
const FOOTNOTE_TABLE_FIELD: u32 = 16;
const TABLE_ENTRIES_FIELD: u32 = 1;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const FOOTNOTE_ANCHOR: char = '\u{000e}';
const FOOTNOTE_ANCHOR_TEXT: &str = "\u{000e}";
const FOOTNOTE_ANCHOR_UNIT: u16 = 0x000e;
const FOOTNOTE_MARK: char = '\u{fffc}';
const FOOTNOTE_CONTENT_PREFIX: &str = "\u{fffc} ";
const FOOTNOTE_REFERENCE_CODEC_RECURSION_LIMIT: u32 = 64;
const MAX_BODY_FOOTNOTES: usize = 4096;

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
        let removed = body_footnote_by_selector(self, selector)?;
        let start = usize::try_from(removed.footnote.position.utf16_index()).map_err(|_| {
            Error::ParseError("Pages footnote position exceeds the platform index range".to_owned())
        })?;
        let end = start
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Pages footnote anchor range overflow".to_owned()))?;
        // Keep the legacy graph edit and its cleanup off the live editor until
        // the removed native reference is absent. Position-only validation is
        // incorrect when a following footnote shifts into the deleted anchor.
        let mut staged = self.clone();
        staged.replace_body_text(start..end, "")?;
        if body_footnote_graphs(&staged, staged.body_storage_id.get())?
            .iter()
            .any(|graph| graph.reference_id == removed.reference_id)
        {
            return Err(Error::InvalidFormat(
                "Pages footnote deletion failed validation".to_owned(),
            ));
        }
        let result = removed.footnote;
        *self = staged;
        Ok(result)
    }
}

pub(super) fn body_footnote_graphs(
    package: &IWorkPackage,
    body_storage_id: u64,
) -> Result<Vec<BodyFootnoteGraph>> {
    let (_, body, body_data) = storage_at_with_data(package, body_storage_id, "Pages body")?;
    let entries = footnote_table_entries(body_storage_id, &body_data, &body)?;
    let mut seen = HashSet::new();
    let limits = footnote_wire_limits(body_data.len())?;
    reserve_footnote_set(
        &mut seen,
        entries.len(),
        limits,
        "Pages body footnote references",
    )?;
    let mut footnotes = Vec::new();
    reserve_footnote_collection(
        &mut footnotes,
        entries.len(),
        limits,
        "Pages body footnote graphs",
    )?;
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
    let remaining_graphs = body_footnote_graphs(package, body_storage_id)?;
    let mut remaining = HashSet::new();
    remaining.try_reserve(remaining_graphs.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "Pages body remaining footnote references",
            amount: remaining_graphs.len(),
        })
    })?;
    remaining.extend(remaining_graphs.into_iter().map(|graph| graph.reference_id));

    let removed_count = before.iter().try_fold(0usize, |count, graph| {
        if remaining.contains(&graph.reference_id) {
            Ok(count)
        } else {
            count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat(
                    "Pages removed footnote graph count overflows usize".to_owned(),
                )
            })
        }
    })?;
    if removed_count == 0 {
        return Ok(());
    }
    let identifier_count = footnote_cleanup_identifier_count(removed_count)?;

    let mut staged = package.clone();
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(identifier_count)
        .map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Pages removed footnote object identifiers",
                amount: identifier_count,
            })
        })?;
    for graph in before
        .iter()
        .filter(|graph| !remaining.contains(&graph.reference_id))
    {
        identifiers.extend(remove_unreferenced_footnote_graph(&mut staged, graph)?);
    }
    release_package_identifier_suffix(&mut staged, &identifiers)?;
    IWorkPackage::from_bytes(&staged.to_bytes()?)?;
    *package = staged;
    Ok(())
}

fn footnote_cleanup_identifier_count(removed_count: usize) -> Result<usize> {
    removed_count.checked_mul(3).ok_or_else(|| {
        Error::InvalidFormat("Pages removed footnote identifier count overflows usize".to_owned())
    })
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
    let reference_data = object_message_data(
        reference_object,
        FOOTNOTE_REFERENCE_MESSAGE_TYPE,
        "Pages footnote reference",
    )?;
    let reference = pages_footnote_codec::decode_footnote_reference(
        reference_data,
        footnote_reference_decode_options(reference_data),
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "Pages footnote reference object {reference_id} failed strict validation: {error}"
        ))
    })?;
    if reference.super_kind().is_some_and(|kind| {
        kind != tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32
    }) {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote object {reference_id} has the wrong attachment kind"
        )));
    }
    let storage_id = reference
        .contained_storage()
        .map(|value| value.identifier().get())
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
    let footnote = Footnote::with_custom_mark(
        position,
        text,
        reference
            .custom_mark_string()
            .map(str::to_owned)
            .map(Into::into),
    )
    .map_err(|error| Error::ParseError(format!("invalid Pages footnote value: {error}")))?;

    Ok(BodyFootnoteGraph {
        footnote,
        reference_id,
        storage_id,
        marker_id,
    })
}

fn footnote_reference_decode_options(source: &[u8]) -> pages_footnote_codec::DecodeOptions {
    pages_footnote_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(16)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        FOOTNOTE_REFERENCE_CODEC_RECURSION_LIMIT,
    )
}

fn footnote_body_attachment_decode_options(source: &[u8]) -> pages_body_codec::DecodeOptions {
    pages_body_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(16)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        FOOTNOTE_REFERENCE_CODEC_RECURSION_LIMIT,
    )
}

fn footnote_wire_limits(source_len: usize) -> Result<WireLimits> {
    WireLimits::default()
        .with_input_bytes(source_len.clamp(1, WireLimits::MAX_INPUT_BYTES))
        .and_then(|limits| limits.with_fields(source_len.clamp(1, WireLimits::MAX_FIELDS)))
        .and_then(|limits| {
            limits.with_rewrite_work(
                source_len
                    .saturating_mul(16)
                    .clamp(1, WireLimits::MAX_REWRITE_WORK),
            )
        })
        .map_err(Into::into)
}

fn reserve_footnote_collection<T>(
    values: &mut Vec<T>,
    additional: usize,
    limits: WireLimits,
    resource: &'static str,
) -> Result<()> {
    let requested = values
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflows usize")))?;
    let limit = limits.max_fields().min(MAX_BODY_FOOTNOTES);
    if requested > limit {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: requested,
            limit,
        }));
    }
    values.try_reserve_exact(additional).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: requested,
        })
    })
}

fn reserve_footnote_set<T: Eq + Hash>(
    values: &mut HashSet<T>,
    additional: usize,
    limits: WireLimits,
    resource: &'static str,
) -> Result<()> {
    let requested = values
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflows usize")))?;
    let limit = limits.max_fields().min(MAX_BODY_FOOTNOTES);
    if requested > limit {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: requested,
            limit,
        }));
    }
    values.try_reserve(additional).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: requested,
        })
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
    let marker_data = object_message_data(
        object,
        TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
        "Pages footnote marker",
    )?;
    let marker = pages_footnote_marker_codec::decode_textual_attachment(
        marker_data,
        footnote_marker_decode_options(marker_data),
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "Pages footnote marker object {marker_id} failed strict validation: {error}"
        ))
    })?;
    if marker.kind() != Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32) {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote marker object {marker_id} has the wrong attachment kind"
        )));
    }
    Ok(())
}

fn footnote_marker_decode_options(source: &[u8]) -> pages_footnote_marker_codec::DecodeOptions {
    pages_footnote_marker_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(16)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        FOOTNOTE_REFERENCE_CODEC_RECURSION_LIMIT,
    )
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
            None => rewrite_repeated_length_delimited_fields(
                &[],
                TABLE_ENTRIES_FIELD,
                &encoded_entries,
            )?,
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
            let entry = pages_body_codec::decode_section_boundary(
                raw,
                footnote_body_attachment_decode_options(raw),
            )
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "Pages body storage {storage_id} footnote attachment failed strict validation: {error}"
                ))
            })?;
            Ok(FootnoteTableEntry {
                index: entry.character_index(),
                reference_id: entry
                    .section()
                    .map_or(0, |value| value.identifier().get()),
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
) -> Result<[u64; 3]> {
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
    Ok([reference_id, graph.storage_id, graph.marker_id])
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
    fn footnote_collection_respects_wire_field_limit_before_reserving() {
        let limits = WireLimits::default().with_fields(1).unwrap();
        let mut graphs = Vec::<()>::new();
        let error =
            reserve_footnote_collection(&mut graphs, 2, limits, "Pages body footnote graphs")
                .unwrap_err();
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: 2,
                limit: 1,
            })
        ));
        assert!(graphs.is_empty());
    }

    #[test]
    fn footnote_cleanup_identifier_count_checks_overflow() {
        let error = footnote_cleanup_identifier_count(usize::MAX).unwrap_err();
        assert!(matches!(
            error,
            Error::InvalidFormat(message)
                if message == "Pages removed footnote identifier count overflows usize"
        ));
    }

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
    fn removing_adjacent_body_footnote_is_atomic_and_tracks_identity() {
        let mut editor = PagesEditor::create_with_text("AB").unwrap();
        let first = editor
            .insert_body_footnote(Position::from_utf16_index(1).unwrap(), "First")
            .unwrap();
        let second = editor
            .insert_body_footnote(Position::from_utf16_index(2).unwrap(), "Second")
            .unwrap();

        let removed = editor
            .remove_body_footnote(Selector::At(first.position))
            .unwrap();

        assert_eq!(removed, first);
        assert_eq!(editor.body_text().unwrap(), "A\u{e}B");
        assert_eq!(
            editor.body_footnotes().unwrap(),
            vec![Footnote {
                position: Position::from_utf16_index(1).unwrap(),
                text: second.text,
                custom_mark: second.custom_mark,
            }]
        );
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
                let data = patch_length_delimited_field(
                    message.data.as_slice(),
                    FOOTNOTE_SUPER_FIELD,
                    false,
                    None,
                )?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: FOOTNOTE_REFERENCE_MESSAGE_TYPE,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let parsed = PagesEditor::from_package(package).unwrap();
        assert_eq!(parsed.body_footnotes().unwrap(), vec![footnote]);
    }

    #[test]
    fn footnote_marker_unknown_fields_survive_an_atomic_text_rewrite() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let graph = body_footnote_graphs(editor.package(), editor.body_storage_id.get())
            .unwrap()
            .pop()
            .unwrap();
        let unknown = [0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'o', b'p', b'a'];
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, graph.marker_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(graph.marker_id).unwrap();
                let message = &object.messages[0];
                let mut data = message.data.clone();
                data.extend_from_slice(&unknown);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut edited = PagesEditor::from_package(package).unwrap();
        assert_eq!(edited.body_footnotes().unwrap()[0].text.as_ref(), "Native");
        edited
            .set_body_footnote_text(Selector::Index(0), "Updated")
            .unwrap();
        let marker_archive = find_object_archive(edited.package(), graph.marker_id).unwrap();
        let marker_archive_data = edited.package().archive(&marker_archive).unwrap();
        let marker = marker_archive_data.object(graph.marker_id).unwrap();
        assert!(marker.messages[0].data.ends_with(&unknown));
    }

    #[test]
    fn malformed_footnote_marker_rewrite_is_failure_atomic() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let marker_id = body_footnote_graphs(editor.package(), editor.body_storage_id.get())
            .unwrap()[0]
            .marker_id;
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, marker_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(marker_id).unwrap();
                let message = &object.messages[0];
                let mut data = message.data.clone();
                data.extend_from_slice(&[0xa0, 0x06, 0x80]);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        let mut malformed = PagesEditor::from_package(package).unwrap();
        let baseline = malformed.to_bytes().unwrap();
        assert!(
            malformed
                .set_body_footnote_text(Selector::Index(0), "Updated")
                .is_err()
        );
        assert_eq!(malformed.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn footnote_reference_unknown_fields_survive_an_atomic_text_rewrite() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let reference_id = body_footnote_graphs(editor.package(), editor.body_storage_id.get())
            .unwrap()[0]
            .reference_id;
        let unknown = [0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'o', b'p', b'a'];
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, reference_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(reference_id).unwrap();
                let message = &object.messages[0];
                let mut data = message.data.clone();
                data.extend_from_slice(&unknown);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut edited = PagesEditor::from_package(package).unwrap();
        assert_eq!(edited.body_footnotes().unwrap()[0].text.as_ref(), "Native");
        edited
            .set_body_footnote_text(Selector::Index(0), "Updated")
            .unwrap();
        let reference_archive = find_object_archive(edited.package(), reference_id).unwrap();
        let reference_archive_data = edited.package().archive(&reference_archive).unwrap();
        let reference = reference_archive_data.object(reference_id).unwrap();
        assert!(reference.messages[0].data.ends_with(&unknown));
    }

    #[test]
    fn malformed_footnote_reference_rewrite_is_failure_atomic() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
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
                let mut data = message.data.clone();
                data.extend_from_slice(&[0xa0, 0x06, 0x80]);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut malformed = PagesEditor::from_package(package).unwrap();
        let baseline = malformed.to_bytes().unwrap();
        assert!(
            malformed
                .set_body_footnote_text(Selector::Index(0), "Updated")
                .is_err()
        );
        assert_eq!(malformed.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn footnote_body_attachment_unknown_fields_survive_an_atomic_text_rewrite() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let body_storage_id = editor.body_storage_id.get();
        let unknown = [0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'o', b'p', b'a'];
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, body_storage_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(body_storage_id).unwrap();
                let message_index = unique_storage_message_index(object, body_storage_id)?;
                let message = &object.messages[message_index];
                let tables = repeated_length_delimited_payloads(
                    message.data.as_slice(),
                    FOOTNOTE_TABLE_FIELD,
                )?;
                let [table] = tables.as_slice() else {
                    return Err(Error::InvalidFormat(
                        "Pages body footnote test requires one attachment table".to_owned(),
                    ));
                };
                let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
                let [entry] = entries.as_slice() else {
                    return Err(Error::InvalidFormat(
                        "Pages body footnote test requires one attachment entry".to_owned(),
                    ));
                };
                let mut replacement = entry.to_vec();
                replacement.extend_from_slice(&unknown);
                let table = rewrite_repeated_length_delimited_fields(
                    table,
                    TABLE_ENTRIES_FIELD,
                    &[replacement],
                )?;
                let data = patch_length_delimited_field(
                    message.data.as_slice(),
                    FOOTNOTE_TABLE_FIELD,
                    true,
                    Some(&table),
                )?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut edited = PagesEditor::from_package(package).unwrap();
        assert_eq!(edited.body_footnotes().unwrap()[0].text.as_ref(), "Native");
        edited
            .set_body_footnote_text(Selector::Index(0), "Updated")
            .unwrap();
        let body_archive = find_object_archive(edited.package(), body_storage_id).unwrap();
        let body_archive_data = edited.package().archive(&body_archive).unwrap();
        let body = body_archive_data.object(body_storage_id).unwrap();
        let message_index = unique_storage_message_index(body, body_storage_id).unwrap();
        let tables = repeated_length_delimited_payloads(
            body.messages[message_index].data.as_slice(),
            FOOTNOTE_TABLE_FIELD,
        )
        .unwrap();
        let [table] = tables.as_slice() else {
            panic!("updated body has one attachment table");
        };
        let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD).unwrap();
        let [entry] = entries.as_slice() else {
            panic!("updated body has one attachment entry");
        };
        assert!(entry.ends_with(&unknown));
    }

    #[test]
    fn malformed_footnote_body_attachment_rewrite_is_failure_atomic() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let body_storage_id = editor.body_storage_id.get();
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, body_storage_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(body_storage_id).unwrap();
                let message_index = unique_storage_message_index(object, body_storage_id)?;
                let message = &object.messages[message_index];
                let tables = repeated_length_delimited_payloads(
                    message.data.as_slice(),
                    FOOTNOTE_TABLE_FIELD,
                )?;
                let [table] = tables.as_slice() else {
                    return Err(Error::InvalidFormat(
                        "Pages body footnote test requires one attachment table".to_owned(),
                    ));
                };
                let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
                let [entry] = entries.as_slice() else {
                    return Err(Error::InvalidFormat(
                        "Pages body footnote test requires one attachment entry".to_owned(),
                    ));
                };
                let mut replacement = entry.to_vec();
                replacement.extend_from_slice(&[0xa0, 0x06, 0x80]);
                let table = rewrite_repeated_length_delimited_fields(
                    table,
                    TABLE_ENTRIES_FIELD,
                    &[replacement],
                )?;
                let data = patch_length_delimited_field(
                    message.data.as_slice(),
                    FOOTNOTE_TABLE_FIELD,
                    true,
                    Some(&table),
                )?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut malformed = PagesEditor::from_package(package).unwrap();
        let baseline = malformed.to_bytes().unwrap();
        assert!(
            malformed
                .set_body_footnote_text(Selector::Index(0), "Updated")
                .is_err()
        );
        assert_eq!(malformed.to_bytes().unwrap(), baseline);
    }
}
