//! Package glue for the W1-D text-engine parity cases (record 0374).
//!
//! The text engines of `litchi-iwa-text-wire` own payloads only; a format
//! package publishes them. No leaf package has a storage-scoped text family
//! yet (W1-B, W1-C and W2 add them), so the leaf route of these cases
//! publishes the engine's payloads through this small, independent glue: it
//! edits the package's archives through the public `litchi_iwa::Edit` and
//! keeps `Index/Metadata.iwa` the way the host does (object UUIDs,
//! cross-component references, the identifier watermark), using the shared
//! `litchi-iwa-common` wire helpers. UUIDs of created objects are derived
//! deterministically, which the harness's `leaf_structural_metadata` policy
//! accepts.

use std::collections::BTreeMap;

use litchi_iwa_archive::iwa::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, UnknownFieldRule,
};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{
    append_repeated_length_delimited_field, remove_repeated_length_delimited_field_where,
    transform_length_delimited_fields_at_path,
};
use litchi_iwa_protos::package_metadata_codec::{
    Batch, RewriteOptions, prepare_package_metadata_rewrite,
};
use litchi_iwa_protos::{tsp, tswp};
use prost::Message;

use crate::harness::TestResult;

pub const METADATA_ENTRY: &str = "Index/Metadata.iwa";
const METADATA_TYPE: u32 = 11_006;
pub const STYLESHEET_TYPE: u32 = 401;
pub const PARAGRAPH_STYLE_TYPE: u32 = 2_022;
const STORAGE_TYPES: [u32; 2] = [2_001, 2_022];

fn error(message: impl Into<String>) -> litchi_iwa::Error {
    litchi_iwa::Error::InvalidFormat(message.into())
}

/// One located message: its archive entry, object and message index.
#[derive(Debug, Clone)]
pub struct Located {
    pub entry: String,
    pub object: u64,
    pub index: usize,
    pub message_type: u32,
    pub payload: Vec<u8>,
}

/// A staged package edit.
pub struct Workspace {
    edit: litchi_iwa::Edit,
}

impl Workspace {
    pub fn open(bytes: &[u8]) -> TestResult<Self> {
        Ok(Self {
            edit: litchi_iwa::Package::from_bytes(bytes)?.edit(),
        })
    }

    pub fn archives(&self) -> TestResult<Vec<(String, Archive)>> {
        let names: Vec<String> = self.edit.iwa_entry_names().map(str::to_owned).collect();
        let mut archives = Vec::new();
        for name in names {
            let archive = self.edit.archive(&name)?;
            archives.push((name, archive));
        }
        Ok(archives)
    }

    /// Every message of `message_type`, by object.
    pub fn messages(&self, message_type: u32) -> TestResult<Vec<Located>> {
        let mut found = Vec::new();
        for (entry, archive) in self.archives()? {
            for object in &archive.objects {
                let Some(identifier) = object.archive_info.identifier else {
                    continue;
                };
                for (index, message) in object.messages.iter().enumerate() {
                    if message.type_ == message_type {
                        found.push(Located {
                            entry: entry.clone(),
                            object: identifier,
                            index,
                            message_type,
                            payload: message.data.clone(),
                        });
                    }
                }
            }
        }
        Ok(found)
    }

    /// The unique message of `message_type` of object `identifier`.
    pub fn message(&self, identifier: u64, message_type: u32) -> TestResult<Located> {
        let found: Vec<Located> = self
            .messages(message_type)?
            .into_iter()
            .filter(|located| located.object == identifier)
            .collect();
        match <[Located; 1]>::try_from(found) {
            Ok([located]) => Ok(located),
            Err(found) => Err(format!(
                "object {identifier} has {} messages of type {message_type}",
                found.len()
            )
            .into()),
        }
    }

    /// The text storage payload of `identifier` (message type 2001, or the
    /// Numbers storage that reuses 2022), as the host locates it.
    pub fn storage(&self, identifier: u64) -> TestResult<Located> {
        for message_type in STORAGE_TYPES {
            for located in self.messages(message_type)? {
                if located.object == identifier
                    && tswp::StorageArchive::decode(located.payload.as_slice()).is_ok()
                {
                    return Ok(located);
                }
            }
        }
        Err(format!("storage {identifier} not found").into())
    }

    /// Every text storage payload.
    pub fn storages(&self) -> TestResult<Vec<Located>> {
        let mut storages = Vec::new();
        for message_type in STORAGE_TYPES {
            for located in self.messages(message_type)? {
                if tswp::StorageArchive::decode(located.payload.as_slice()).is_ok() {
                    storages.push(located);
                }
            }
        }
        Ok(storages)
    }

    pub fn update(
        &mut self,
        entry: &str,
        update: impl FnOnce(&mut Archive) -> litchi_iwa::Result<()>,
    ) -> TestResult {
        self.edit.update_archive(entry, update)?;
        Ok(())
    }

    /// Replace one message payload, rewriting the object references of its
    /// metadata from `from` to `to` (exactly one occurrence) when given.
    pub fn replace(
        &mut self,
        located: &Located,
        payload: Vec<u8>,
        repoint: Option<(u64, u64)>,
    ) -> TestResult {
        let (object, index, message_type) = (located.object, located.index, located.message_type);
        self.update(&located.entry, |archive| {
            let target = archive
                .object_mut(object)
                .ok_or_else(|| error("object disappeared"))?;
            target.replace_message(
                index,
                RawMessage {
                    type_: message_type,
                    data: payload,
                },
                Default::default(),
            )?;
            if let Some((from, to)) = repoint {
                let info = &mut target.archive_info.message_infos.as_mut_slice()[index];
                let mut replaced = 0;
                for reference in &mut info.object_references {
                    if *reference == from {
                        *reference = to;
                        replaced += 1;
                    }
                }
                for field in &mut info.field_infos {
                    for reference in &mut field.object_references {
                        if *reference == from {
                            *reference = to;
                        }
                    }
                }
                if replaced != 1 {
                    return Err(error(format!("{replaced} references to {from}")));
                }
            }
            Ok(())
        })
    }

    /// The host's next object identifier: one above every archive object,
    /// the metadata watermark and every identifier the metadata names.
    pub fn next_identifier(&self) -> TestResult<u64> {
        let mut maximum = 0u64;
        for (_, archive) in self.archives()? {
            for object in &archive.objects {
                maximum = maximum.max(object.archive_info.identifier.unwrap_or_default());
            }
        }
        if let Some(metadata) = self.metadata()? {
            maximum = maximum.max(metadata.last_object_identifier);
            for component in metadata
                .components
                .iter()
                .chain(&metadata.versioned_components)
            {
                maximum = maximum.max(component.identifier);
                for identifier in &component.ambiguous_object_identifiers {
                    maximum = maximum.max(*identifier);
                }
                for entry in &component.object_uuid_map_entries {
                    maximum = maximum.max(entry.identifier);
                }
                for reference in &component.external_references {
                    maximum = maximum.max(reference.object_identifier.unwrap_or_default());
                }
                for owner in &component.data_references {
                    for entry in &owner.object_reference_list {
                        maximum = maximum.max(entry.object_identifier);
                    }
                }
            }
        }
        Ok(maximum + 1)
    }

    fn metadata_payload(&self) -> TestResult<Option<Vec<u8>>> {
        if !self.edit.contains_entry(METADATA_ENTRY) {
            return Ok(None);
        }
        let archive = self.edit.archive(METADATA_ENTRY)?;
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ == METADATA_TYPE {
                    return Ok(Some(message.data.clone()));
                }
            }
        }
        Ok(None)
    }

    pub fn metadata(&self) -> TestResult<Option<tsp::PackageMetadata>> {
        Ok(self
            .metadata_payload()?
            .map(|payload| tsp::PackageMetadata::decode(payload.as_slice()))
            .transpose()?)
    }

    fn rewrite_metadata(
        &mut self,
        rewrite: impl FnOnce(&[u8]) -> litchi_iwa::Result<Vec<u8>>,
    ) -> TestResult {
        if !self.edit.contains_entry(METADATA_ENTRY) {
            return Ok(());
        }
        self.update(METADATA_ENTRY, |archive| {
            let object = archive
                .objects
                .iter_mut()
                .find(|object| object.messages.iter().any(|m| m.type_ == METADATA_TYPE))
                .ok_or_else(|| error("no package metadata"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == METADATA_TYPE)
                .ok_or_else(|| error("no package metadata"))?;
            let data = rewrite(&object.messages[index].data)?;
            object.replace_message(
                index,
                RawMessage {
                    type_: METADATA_TYPE,
                    data,
                },
                Default::default(),
            )?;
            Ok(())
        })
    }

    /// The metadata component of an archive entry (`Index/<locator>.iwa`).
    pub fn component(&self, entry: &str) -> TestResult<Option<u64>> {
        let locator = entry
            .strip_prefix("Index/")
            .and_then(|name| name.strip_suffix(".iwa"))
            .ok_or("invalid IWA component name")?;
        let Some(metadata) = self.metadata()? else {
            return Ok(None);
        };
        let matches: Vec<u64> = metadata
            .components
            .iter()
            .filter(|component| {
                component
                    .locator
                    .as_deref()
                    .unwrap_or(&component.preferred_locator)
                    == locator
            })
            .map(|component| component.identifier)
            .collect();
        match matches.as_slice() {
            [] => Ok(None),
            [component] => Ok(Some(*component)),
            _ => Err("several components share one locator".into()),
        }
    }

    fn per_component(
        &mut self,
        component: u64,
        mut transform: impl FnMut(&[u8]) -> litchi_iwa::Result<Vec<u8>>,
    ) -> TestResult {
        self.rewrite_metadata(|data| {
            transform_length_delimited_fields_at_path(data, &[3], |component_data| {
                let decoded = tsp::ComponentInfo::decode(component_data)?;
                if decoded.identifier == component {
                    transform(component_data)
                } else {
                    Ok(component_data.to_vec())
                }
            })
        })
    }

    pub fn add_object_uuid(&mut self, component: u64, identifier: u64) -> TestResult {
        let entry = tsp::ObjectUuidMapEntry {
            identifier,
            uuid: tsp::Uuid {
                lower: 0x4c49_5443_4849_0000 ^ identifier,
                upper: 0x5731_442d_7465_7874,
            },
        }
        .encode_to_vec();
        self.per_component(component, |data| {
            Ok(append_repeated_length_delimited_field(data, 11, &entry)?)
        })
    }

    pub fn remove_object_uuid(&mut self, component: u64, identifier: u64) -> TestResult {
        self.per_component(component, |data| {
            remove_repeated_length_delimited_field_where(data, 11, |entry| {
                Ok(tsp::ObjectUuidMapEntry::decode(entry)?.identifier == identifier)
            })
        })
    }

    /// Add a `source → target` reference to `object` unless present.
    pub fn add_external_reference(&mut self, source: u64, target: u64, object: u64) -> TestResult {
        let reference = tsp::ComponentExternalReference {
            component_identifier: target,
            object_identifier: Some(object),
            is_weak: None,
        };
        self.per_component(source, |data| {
            let component = tsp::ComponentInfo::decode(data)?;
            if component.external_references.contains(&reference) {
                return Ok(data.to_vec());
            }
            Ok(append_repeated_length_delimited_field(
                data,
                6,
                &reference.encode_to_vec(),
            )?)
        })
    }

    /// Remove every reference to `object` in `target` from every component.
    pub fn remove_external_references_to(&mut self, target: u64, object: u64) -> TestResult {
        self.rewrite_metadata(|data| {
            transform_length_delimited_fields_at_path(data, &[3], |component_data| {
                let component = tsp::ComponentInfo::decode(component_data)?;
                if !component.external_references.iter().any(|reference| {
                    reference.component_identifier == target
                        && reference.object_identifier == Some(object)
                }) {
                    return Ok(component_data.to_vec());
                }
                remove_repeated_length_delimited_field_where(component_data, 6, |payload| {
                    let reference = tsp::ComponentExternalReference::decode(payload)?;
                    Ok(reference.component_identifier == target
                        && reference.object_identifier == Some(object))
                })
            })
        })
    }

    /// Raise the watermark to `identifier` (the host's `Batch` rewrite).
    pub fn raise_watermark(&mut self, identifier: u64) -> TestResult {
        let Some(metadata) = self.metadata()? else {
            return Ok(());
        };
        let current = metadata.last_object_identifier;
        if identifier <= current {
            return Ok(());
        }
        let options = RewriteOptions::new(
            WireLimits::MAX_INPUT_BYTES,
            WireLimits::MAX_OUTPUT_BYTES,
            WireLimits::MAX_FIELDS,
            WireLimits::MAX_REWRITE_WORK,
            64,
            WireLimits::MAX_FIELDS,
            WireLimits::MAX_FIELDS,
            0,
        );
        self.rewrite_metadata(|data| {
            let prepared = prepare_package_metadata_rewrite(
                data,
                Batch::new(current, identifier, &[], &[]),
                options,
            )
            .map_err(|e| error(format!("{e:?}")))?;
            let limits = prepared.execution_requirements().exact_limits();
            Ok(prepared
                .execute(limits)
                .map_err(|e| error(format!("{e:?}")))?
                .into_bytes())
        })
    }

    /// Lower the watermark after removing `removed` when it named the
    /// highest object (the host's `release_package_identifier_suffix`).
    pub fn release_watermark(&mut self, removed: &[u64]) -> TestResult {
        let Some(metadata) = self.metadata()? else {
            return Ok(());
        };
        let last = metadata.last_object_identifier;
        if !removed.contains(&last) {
            return Ok(());
        }
        let mut maximum = 0u64;
        for (_, archive) in self.archives()? {
            for object in &archive.objects {
                maximum = maximum.max(object.archive_info.identifier.unwrap_or_default());
            }
        }
        if maximum == last {
            return Err("released watermark object remains".into());
        }
        self.rewrite_metadata(|data| {
            Ok(litchi_iwa_common::wire::patch_varint_field(
                data,
                1,
                true,
                Some(maximum),
            )?)
        })
    }

    pub fn into_bytes(self) -> TestResult<Vec<u8>> {
        let commit = self.edit.commit()?;
        Ok(commit.package().to_bytes()?)
    }
}

/// The host's object-reference count of `style` in storage paragraph-style
/// tables and whether any paragraph style names it as parent.
pub fn exclusive(workspace: &Workspace, style: u64) -> TestResult<bool> {
    let mut references = 0usize;
    for storage in workspace.storages()? {
        let decoded = tswp::StorageArchive::decode(storage.payload.as_slice())?;
        references += decoded
            .table_para_style
            .iter()
            .flat_map(|table| &table.entries)
            .filter(|entry| entry.object.as_ref().is_some_and(|r| r.identifier == style))
            .count();
    }
    for located in workspace.messages(PARAGRAPH_STYLE_TYPE)? {
        if let Ok(decoded) = tswp::ParagraphStyleArchive::decode(located.payload.as_slice())
            && decoded
                .super_
                .parent
                .is_some_and(|parent| parent.identifier == style)
        {
            return Ok(false);
        }
    }
    Ok(references == 1)
}

/// Every paragraph-style payload by identifier, with its archive entry.
pub fn paragraph_styles(
    workspace: &Workspace,
) -> TestResult<(BTreeMap<u64, Vec<u8>>, BTreeMap<u64, String>)> {
    let mut styles = BTreeMap::new();
    let mut entries = BTreeMap::new();
    for located in workspace.messages(PARAGRAPH_STYLE_TYPE)? {
        if tswp::ParagraphStyleArchive::decode(located.payload.as_slice()).is_ok() {
            entries.insert(located.object, located.entry.clone());
            styles.insert(located.object, located.payload);
        }
    }
    Ok((styles, entries))
}

/// The managed field-path metadata of a paragraph-style variation.
pub fn managed_field_info(path: [u32; 2]) -> FieldInfo {
    FieldInfo {
        path: FieldPath {
            path: path.to_vec(),
        },
        r#type: Some(FieldType::Message),
        unknown_field_rule: Some(UnknownFieldRule::IgnoreAndPreserve),
        ..Default::default()
    }
}

/// The host's `sync_managed_field_info`.
pub fn sync_field_info(infos: &mut Vec<FieldInfo>, present: bool, path: [u32; 2]) {
    let managed = |field: &FieldInfo| field.path.path == path;
    if present {
        if !infos.iter().any(managed) {
            infos.push(managed_field_info(path));
        }
    } else {
        infos.retain(|field| !managed(field));
    }
}

/// A new object holding one message.
pub fn new_object(
    identifier: u64,
    message_type: u32,
    payload: Vec<u8>,
    version: &[u32],
    references: &[u64],
    field_infos: Vec<FieldInfo>,
) -> TestResult<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: payload,
        }],
        Default::default(),
    )?;
    let info = &mut object.archive_info.message_infos.as_mut_slice()[0];
    info.versions = version.to_vec().into();
    info.object_references = references.to_vec();
    info.field_infos = field_infos;
    Ok(object)
}
