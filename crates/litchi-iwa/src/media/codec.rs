//! Media signatures, catalog validation, and iWork metadata wire edits.

use std::collections::{HashMap, HashSet, VecDeque};
use std::io::Write;
use std::ops::Range;
use std::path::{Component, Path};

use prost::Message;

use crate::package::IWorkPackage;
use crate::protobuf;
use litchi_iwa_common::varint::{decode_varint_from_bytes, encode_varint_into};
use litchi_iwa_graph::ObjectId;
use crate::{Error, Result};

use super::model::{EmbeddedMediaAsset, MediaAsset, MediaAssetId, MediaLimits, MediaType};

pub(crate) const PACKAGE_METADATA_ENTRY: &str = "Index/Metadata.iwa";
pub(crate) const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const DATA_METADATA_MAP_MESSAGE_TYPE: u32 = 11_015;

pub(crate) fn insert_unique_asset(
    assets: &mut HashMap<String, MediaAsset>,
    asset: MediaAsset,
    limits: MediaLimits,
    total_size: &mut u64,
) -> Result<()> {
    if let Some(previous) = assets.get(&asset.filename) {
        return Err(Error::Bundle(format!(
            "Media basenames are ambiguous: {} and {}",
            previous.path.display(),
            asset.path.display()
        )));
    }
    if assets.len() >= limits.max_assets {
        return Err(Error::Bundle(format!(
            "Media asset count exceeds the configured {}-entry limit",
            limits.max_assets
        )));
    }
    if asset.size > limits.max_asset_bytes {
        return Err(Error::Bundle(format!(
            "Media asset {} is {} bytes, exceeding the configured {}-byte limit",
            asset.path.display(),
            asset.size,
            limits.max_asset_bytes
        )));
    }
    let new_total = total_size
        .checked_add(asset.size)
        .ok_or_else(|| Error::Bundle("Aggregate media size overflows u64".to_owned()))?;
    if new_total > limits.max_total_bytes {
        return Err(Error::Bundle(format!(
            "Aggregate media size exceeds the configured {}-byte limit",
            limits.max_total_bytes
        )));
    }
    assets.insert(asset.filename.clone(), asset);
    *total_size = new_total;
    Ok(())
}

pub(crate) fn write_package_entry<W: Write>(
    package: &IWorkPackage,
    asset: &MediaAsset,
    limits: MediaLimits,
    sink: &mut W,
) -> Result<()> {
    let name = asset.path.to_str().ok_or_else(|| {
        Error::Bundle(format!(
            "Media path is not valid UTF-8: {}",
            asset.path.display()
        ))
    })?;
    let data = package
        .entry(name)
        .ok_or_else(|| Error::Bundle(format!("Media package entry not found: {name}")))?;
    if u64::try_from(data.len()).unwrap_or(u64::MAX) > limits.max_asset_bytes {
        return Err(Error::Bundle(format!(
            "Media package entry {name} exceeds the configured {}-byte limit",
            limits.max_asset_bytes
        )));
    }
    sink.write_all(data)?;
    Ok(())
}

pub(crate) fn validate_replacement_type(
    asset: &EmbeddedMediaAsset,
    replacement: &[u8],
) -> Result<()> {
    let detected = MediaType::from_bytes(replacement);
    if asset.media_type != MediaType::Unknown
        && detected != MediaType::Unknown
        && asset.media_type != detected
    {
        return Err(Error::Bundle(format!(
            "Replacement signature is {}, but {} is declared as {}",
            detected.name(),
            asset.preferred_filename,
            asset.media_type.name()
        )));
    }
    Ok(())
}

pub(crate) fn validate_new_media(filename: &str, data: &[u8], maximum_length: usize) -> Result<()> {
    if data.is_empty() {
        return Err(Error::Bundle(
            "A materialized media asset cannot contain empty data".to_owned(),
        ));
    }
    if data.len() > maximum_length {
        return Err(Error::Bundle(format!(
            "Media is {} bytes, exceeding the configured {}-byte limit",
            data.len(),
            maximum_length
        )));
    }
    let path = Path::new(filename);
    if path.file_name().and_then(|name| name.to_str()) != Some(filename) {
        return Err(Error::Bundle(format!(
            "Preferred media filename must be a safe basename: {filename:?}"
        )));
    }
    data_entry_name(filename)?;
    let expected = path
        .extension()
        .and_then(|extension| extension.to_str())
        .map(MediaType::from_extension)
        .unwrap_or(MediaType::Unknown);
    let detected = MediaType::from_bytes(data);
    if expected != MediaType::Unknown && detected != MediaType::Unknown && expected != detected {
        return Err(Error::Bundle(format!(
            "Media signature is {}, but {filename} is declared as {}",
            detected.name(),
            expected.name()
        )));
    }
    Ok(())
}

pub(crate) fn materialized_file_name(
    preferred_filename: &str,
    data_identifier: MediaAssetId,
) -> Result<String> {
    let path = Path::new(preferred_filename);
    let stem = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.is_empty())
        .ok_or_else(|| Error::Bundle("Preferred media filename has no stem".to_owned()))?;
    Ok(
        match path.extension().and_then(|extension| extension.to_str()) {
            Some(extension) if !extension.is_empty() => {
                format!("{stem}-{data_identifier}.{extension}")
            },
            _ => format!("{stem}-{data_identifier}"),
        },
    )
}

pub(crate) fn embedded_assets(package: &IWorkPackage) -> Result<Vec<EmbeddedMediaAsset>> {
    let metadata = decode_package_metadata(package)?;
    let mut component_counts = HashMap::<u64, u64>::new();
    let mut component_record_counts = HashMap::<u64, usize>::new();
    let mut referencing_objects = HashMap::<u64, HashSet<ObjectId>>::new();
    for component in metadata
        .components
        .iter()
        .chain(metadata.versioned_components.iter())
    {
        for reference in &component.data_references {
            let data_identifier = MediaAssetId::try_from(reference.data_identifier)?.get();
            let record_count = component_record_counts.entry(data_identifier).or_default();
            *record_count = record_count.checked_add(1).ok_or_else(|| {
                Error::Bundle("Component data reference record count overflow".to_owned())
            })?;
            let count = reference
                .object_reference_list
                .iter()
                .try_fold(0u64, |sum, object| sum.checked_add(u64::from(object.count)))
                .ok_or_else(|| {
                    Error::Bundle("Component data reference count overflow".to_owned())
                })?;
            let current = component_counts.entry(data_identifier).or_default();
            *current = current.checked_add(count).ok_or_else(|| {
                Error::Bundle("Component data reference count overflow".to_owned())
            })?;
            for object in &reference.object_reference_list {
                let object_identifier = ObjectId::try_from(object.object_identifier).map_err(
                    |_| {
                        Error::InvalidFormat(format!(
                            "Component data reference for media {data_identifier} contains a zero object identifier"
                        ))
                    },
                )?;
                referencing_objects
                    .entry(data_identifier)
                    .or_default()
                    .insert(object_identifier);
            }
        }
    }

    let mut message_counts = HashMap::<u64, usize>::new();
    let metadata_map_identifier = metadata
        .data_metadata_map
        .as_ref()
        .map(|reference| {
            ObjectId::try_from(reference.identifier).map_err(|_| {
                Error::InvalidFormat(
                    "DataMetadataMap reference contains a zero object identifier".to_owned(),
                )
            })
        })
        .transpose()?;
    let mut data_metadata_ids = HashSet::new();
    let mut metadata_map_payloads = 0usize;
    let iwa_names = package
        .iwa_entry_names()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    for name in iwa_names {
        let archive = package.archive(&name)?;
        for object in archive.objects {
            let object_identifier = object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat(format!("Object in {name} has no identifier"))
            })?;
            let object_identifier = ObjectId::try_from(object_identifier).map_err(|_| {
                Error::InvalidFormat(format!("Object in {name} has a zero archive identifier"))
            })?;
            if Some(object_identifier) == metadata_map_identifier {
                for message in &object.messages {
                    if message.type_ == DATA_METADATA_MAP_MESSAGE_TYPE {
                        metadata_map_payloads =
                            metadata_map_payloads.checked_add(1).ok_or_else(|| {
                                Error::Bundle("DataMetadataMap payload count overflow".to_owned())
                            })?;
                        let map = protobuf::tsp::DataMetadataMap::decode(message.data.as_slice())?;
                        for entry in map.data_metadata_entries {
                            let data_identifier = MediaAssetId::try_from(entry.data_identifier)?;
                            data_metadata_ids.insert(data_identifier.get());
                        }
                    }
                }
            }
            for info in object.archive_info.message_infos {
                for identifier in info.data_references {
                    let data_identifier = MediaAssetId::try_from(identifier)?.get();
                    let count = message_counts.entry(data_identifier).or_default();
                    *count = count.checked_add(1).ok_or_else(|| {
                        Error::Bundle("Message data reference count overflow".to_owned())
                    })?;
                    referencing_objects
                        .entry(data_identifier)
                        .or_default()
                        .insert(object_identifier);
                }
            }
        }
    }
    if metadata_map_identifier.is_some() && metadata_map_payloads != 1 {
        return Err(Error::Bundle(format!(
            "Expected one DataMetadataMap payload, found {metadata_map_payloads}"
        )));
    }

    let mut assets = Vec::with_capacity(metadata.datas.len());
    let mut identifiers = std::collections::HashSet::with_capacity(metadata.datas.len());
    for data in metadata.datas {
        let data_identifier = MediaAssetId::try_from(data.identifier)?;
        if !identifiers.insert(data.identifier) {
            return Err(Error::Bundle(format!(
                "Duplicate DataInfo identifier {}",
                data.identifier
            )));
        }
        let package_path = data
            .file_name
            .as_deref()
            .filter(|file_name| !file_name.is_empty())
            .map(data_entry_name)
            .transpose()?
            .filter(|path| package.contains_entry(path));
        let size = package_path
            .as_deref()
            .and_then(|path| package.entry(path))
            .map(|bytes| u64::try_from(bytes.len()))
            .transpose()
            .map_err(|_| Error::Bundle("Materialized asset length exceeds u64".to_owned()))?;
        let type_name = package_path
            .as_deref()
            .and_then(|path| Path::new(path).file_name())
            .and_then(|name| name.to_str())
            .unwrap_or(&data.preferred_file_name);
        let media_type = Path::new(type_name)
            .extension()
            .and_then(|extension| extension.to_str())
            .map(MediaType::from_extension)
            .unwrap_or(MediaType::Unknown);
        assets.push(EmbeddedMediaAsset {
            data_identifier,
            preferred_filename: data.preferred_file_name,
            package_path,
            media_type,
            size,
            declared_size: data.materialized_length,
            digest: data.digest,
            component_reference_count: component_counts.get(&data.identifier).copied().unwrap_or(0),
            component_reference_record_count: component_record_counts
                .get(&data.identifier)
                .copied()
                .unwrap_or(0),
            message_reference_count: message_counts.get(&data.identifier).copied().unwrap_or(0),
            has_data_metadata: data_metadata_ids.contains(&data.identifier),
            referencing_object_ids: {
                let mut identifiers = referencing_objects
                    .remove(&data.identifier)
                    .unwrap_or_default()
                    .into_iter()
                    .collect::<Vec<_>>();
                identifiers.sort_unstable();
                identifiers
            },
        });
    }
    assets.sort_unstable_by_key(|asset| asset.data_identifier);
    Ok(assets)
}

pub(crate) fn reachable_embedded_assets(
    package: &IWorkPackage,
    roots: impl IntoIterator<Item = u64>,
) -> Result<Vec<EmbeddedMediaAsset>> {
    let assets = embedded_assets(package)?;
    let mut outgoing = HashMap::<ObjectId, Vec<ObjectId>>::new();
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        for object in archive.objects {
            let identifier = object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat(format!("Object in {name} has no identifier"))
            })?;
            let identifier = ObjectId::try_from(identifier).map_err(|_| {
                Error::InvalidFormat(format!("Object in {name} has a zero archive identifier"))
            })?;
            let references = outgoing.entry(identifier).or_default();
            for info in object.archive_info.message_infos {
                for reference in info.object_references {
                    let reference = ObjectId::try_from(reference).map_err(|_| {
                        Error::InvalidFormat(format!(
                            "Object {} in {name} contains a zero object reference",
                            identifier.get()
                        ))
                    })?;
                    if !references.contains(&reference) {
                        references.push(reference);
                    }
                }
            }
        }
    }

    let mut reachable = HashSet::<ObjectId>::new();
    let mut queue = roots
        .into_iter()
        .map(|root| {
            ObjectId::try_from(root).map_err(|_| {
                Error::InvalidFormat("Media reachability root must be non-zero".to_owned())
            })
        })
        .collect::<Result<VecDeque<_>>>()?;
    while let Some(identifier) = queue.pop_front() {
        if !reachable.insert(identifier) {
            continue;
        }
        if let Some(references) = outgoing.get(&identifier) {
            queue.extend(references.iter().copied());
        }
    }
    Ok(assets
        .into_iter()
        .filter(|asset| {
            asset
                .referencing_object_ids
                .iter()
                .any(|identifier| reachable.contains(identifier))
        })
        .collect())
}

fn decode_package_metadata(package: &IWorkPackage) -> Result<protobuf::tsp::PackageMetadata> {
    let archive = package.archive(PACKAGE_METADATA_ENTRY)?;
    let mut payload = None;
    for object in &archive.objects {
        for message in &object.messages {
            if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                && payload.replace(message.data.as_slice()).is_some()
            {
                return Err(Error::Bundle(
                    "Package contains multiple PackageMetadata payloads".to_owned(),
                ));
            }
        }
    }
    protobuf::tsp::PackageMetadata::decode(
        payload.ok_or_else(|| Error::Bundle("PackageMetadata payload was not found".to_owned()))?,
    )
    .map_err(Into::into)
}

pub(crate) fn data_entry_name(file_name: &str) -> Result<String> {
    if file_name.is_empty() || file_name.contains(['\0', '\\']) {
        return Err(Error::Bundle(format!(
            "Unsafe DataInfo filename: {file_name:?}"
        )));
    }
    let path = Path::new(file_name);
    if path.is_absolute()
        || path
            .components()
            .any(|component| !matches!(component, Component::Normal(_)))
    {
        return Err(Error::Bundle(format!(
            "Unsafe DataInfo filename: {file_name:?}"
        )));
    }
    Ok(format!("Data/{file_name}"))
}

#[derive(Debug, Clone)]
pub(crate) struct WireField {
    pub(crate) number: u32,
    pub(crate) wire_type: u8,
    pub(crate) start: usize,
    pub(crate) key_end: usize,
    pub(crate) end: usize,
    pub(crate) payload: Option<Range<usize>>,
}

pub(crate) fn parse_wire_fields(data: &[u8]) -> Result<Vec<WireField>> {
    let mut fields = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let start = offset;
        let (key, key_length) = decode_varint_from_bytes(&data[offset..])
            .map_err(|error| Error::InvalidFormat(format!("Invalid protobuf key: {error}")))?;
        offset = offset
            .checked_add(key_length)
            .ok_or_else(|| Error::InvalidFormat("Protobuf key offset overflow".to_owned()))?;
        let number = key >> 3;
        if number == 0 || number > 0x1fff_ffff {
            return Err(Error::InvalidFormat(format!(
                "Invalid protobuf field number {number}"
            )));
        }
        let wire_type = (key & 7) as u8;
        let key_end = offset;
        let payload = match wire_type {
            0 => {
                let (_, length) = decode_varint_from_bytes(&data[offset..]).map_err(|error| {
                    Error::InvalidFormat(format!("Invalid protobuf varint value: {error}"))
                })?;
                offset = offset.checked_add(length).ok_or_else(|| {
                    Error::InvalidFormat("Protobuf varint offset overflow".to_owned())
                })?;
                None
            },
            1 => {
                offset = offset.checked_add(8).ok_or_else(|| {
                    Error::InvalidFormat("Protobuf fixed64 offset overflow".to_owned())
                })?;
                None
            },
            2 => {
                let (length, prefix_length) =
                    decode_varint_from_bytes(&data[offset..]).map_err(|error| {
                        Error::InvalidFormat(format!("Invalid protobuf length: {error}"))
                    })?;
                offset = offset.checked_add(prefix_length).ok_or_else(|| {
                    Error::InvalidFormat("Protobuf length prefix overflow".to_owned())
                })?;
                let payload_start = offset;
                let length = usize::try_from(length).map_err(|_| {
                    Error::InvalidFormat("Protobuf field length exceeds usize".to_owned())
                })?;
                offset = offset.checked_add(length).ok_or_else(|| {
                    Error::InvalidFormat("Protobuf field range overflow".to_owned())
                })?;
                Some(payload_start..offset)
            },
            5 => {
                offset = offset.checked_add(4).ok_or_else(|| {
                    Error::InvalidFormat("Protobuf fixed32 offset overflow".to_owned())
                })?;
                None
            },
            3 | 4 => {
                return Err(Error::InvalidFormat(
                    "Deprecated protobuf groups are not supported in PackageMetadata".to_owned(),
                ));
            },
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "Invalid protobuf wire type {wire_type}"
                )));
            },
        };
        if offset > data.len() {
            return Err(Error::InvalidFormat(
                "Truncated protobuf field in PackageMetadata".to_owned(),
            ));
        }
        fields.push(WireField {
            number: number as u32,
            wire_type,
            start,
            key_end,
            end: offset,
            payload,
        });
    }
    Ok(fields)
}

pub(crate) fn field_payload<'a>(data: &'a [u8], field: &WireField) -> Result<&'a [u8]> {
    let range = field.payload.clone().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Protobuf field {} is not length-delimited",
            field.number
        ))
    })?;
    data.get(range)
        .ok_or_else(|| Error::InvalidFormat("Protobuf payload range is invalid".to_owned()))
}

fn field_varint(data: &[u8], field: &WireField) -> Result<u64> {
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "Protobuf field {} is not a varint",
            field.number
        )));
    }
    decode_varint_from_bytes(
        data.get(field.key_end..field.end)
            .ok_or_else(|| Error::InvalidFormat("Protobuf varint range is invalid".to_owned()))?,
    )
    .map(|(value, _)| value)
    .map_err(|error| Error::InvalidFormat(format!("Invalid protobuf varint: {error}")))
}

fn data_info_identifier(data: &[u8]) -> Result<u64> {
    let fields = parse_wire_fields(data)?;
    let identifiers = fields
        .iter()
        .filter(|field| field.number == 1)
        .map(|field| field_varint(data, field))
        .collect::<Result<Vec<_>>>()?;
    match identifiers.as_slice() {
        [identifier] => Ok(*identifier),
        [] => Err(Error::InvalidFormat(
            "DataInfo is missing its required identifier".to_owned(),
        )),
        _ => Err(Error::InvalidFormat(
            "DataInfo contains duplicate identifiers".to_owned(),
        )),
    }
}

pub(crate) fn patch_package_metadata(
    metadata: &[u8],
    data_identifier: u64,
    digest: &[u8],
    materialized_length: u64,
) -> Result<Vec<u8>> {
    if digest.len() != 20 {
        return Err(Error::InvalidFormat(format!(
            "iWork materialized data digest must be SHA-1 (20 bytes), got {}",
            digest.len()
        )));
    }
    let fields = parse_wire_fields(metadata)?;
    let mut output = Vec::with_capacity(metadata.len());
    let mut patched_count = 0usize;
    for field in fields {
        if field.number == 4 {
            if field.wire_type != 2 {
                return Err(Error::InvalidFormat(
                    "PackageMetadata.datas has an invalid wire type".to_owned(),
                ));
            }
            let data_info = field_payload(metadata, &field)?;
            if data_info_identifier(data_info)? == data_identifier {
                patched_count = patched_count.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Patched DataInfo count overflow".to_owned())
                })?;
                let patched = patch_data_info(data_info, digest, materialized_length)?;
                output.extend_from_slice(&metadata[field.start..field.key_end]);
                encode_varint_into(&mut output, patched.len() as u64);
                output.extend_from_slice(&patched);
                continue;
            }
        }
        output.extend_from_slice(&metadata[field.start..field.end]);
    }
    match patched_count {
        1 => {},
        0 => {
            return Err(Error::Bundle(format!(
                "Data identifier {data_identifier} is absent from PackageMetadata"
            )));
        },
        _ => {
            return Err(Error::Bundle(format!(
                "Data identifier {data_identifier} is duplicated in PackageMetadata"
            )));
        },
    }
    let decoded = protobuf::tsp::PackageMetadata::decode(output.as_slice())?;
    let matches = decoded
        .datas
        .iter()
        .filter(|data| data.identifier == data_identifier)
        .collect::<Vec<_>>();
    if matches.len() != 1
        || matches[0].digest != digest
        || matches[0].materialized_length != Some(materialized_length)
    {
        return Err(Error::InvalidFormat(
            "Patched PackageMetadata did not decode to the requested values".to_owned(),
        ));
    }
    Ok(output)
}

pub(crate) fn append_data_info(
    metadata: &[u8],
    data_identifier: u64,
    digest: &[u8],
    preferred_filename: &str,
    file_name: &str,
    materialized_length: u64,
) -> Result<Vec<u8>> {
    if digest.len() != 20 {
        return Err(Error::InvalidFormat(format!(
            "iWork materialized data digest must be SHA-1 (20 bytes), got {}",
            digest.len()
        )));
    }
    let decoded = protobuf::tsp::PackageMetadata::decode(metadata)?;
    if decoded
        .datas
        .iter()
        .any(|data| data.identifier == data_identifier)
    {
        return Err(Error::Bundle(format!(
            "Data identifier {data_identifier} already exists"
        )));
    }

    let mut data_info = Vec::new();
    append_wire_varint(&mut data_info, 1, data_identifier);
    append_wire_bytes(&mut data_info, 2, digest);
    append_wire_bytes(&mut data_info, 3, preferred_filename.as_bytes());
    append_wire_bytes(&mut data_info, 4, file_name.as_bytes());
    append_wire_varint(&mut data_info, 18, materialized_length);

    // Appending a repeated field is protobuf-canonical and avoids rewriting any
    // pre-existing metadata field, including unknown extensions.
    let mut output = Vec::with_capacity(metadata.len() + data_info.len() + 16);
    output.extend_from_slice(metadata);
    append_wire_bytes(&mut output, 4, &data_info);
    let verified = protobuf::tsp::PackageMetadata::decode(output.as_slice())?;
    let inserted = verified
        .datas
        .iter()
        .filter(|data| data.identifier == data_identifier)
        .collect::<Vec<_>>();
    if inserted.len() != 1
        || inserted[0].digest != digest
        || inserted[0].preferred_file_name != preferred_filename
        || inserted[0].file_name.as_deref() != Some(file_name)
        || inserted[0].materialized_length != Some(materialized_length)
    {
        return Err(Error::InvalidFormat(
            "Appended DataInfo did not decode to the requested values".to_owned(),
        ));
    }
    Ok(output)
}

fn append_wire_varint(output: &mut Vec<u8>, field_number: u64, value: u64) {
    encode_varint_into(output, field_number << 3);
    encode_varint_into(output, value);
}

fn append_wire_bytes(output: &mut Vec<u8>, field_number: u64, value: &[u8]) {
    encode_varint_into(output, (field_number << 3) | 2);
    encode_varint_into(output, value.len() as u64);
    output.extend_from_slice(value);
}

pub(crate) fn remove_data_info(metadata: &[u8], data_identifier: u64) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(metadata)?;
    let mut output = Vec::with_capacity(metadata.len());
    let mut removed_count = 0usize;
    for field in fields {
        if field.number == 4 {
            if field.wire_type != 2 {
                return Err(Error::InvalidFormat(
                    "PackageMetadata.datas has an invalid wire type".to_owned(),
                ));
            }
            if data_info_identifier(field_payload(metadata, &field)?)? == data_identifier {
                removed_count = removed_count.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Removed DataInfo count overflow".to_owned())
                })?;
                continue;
            }
        }
        output.extend_from_slice(&metadata[field.start..field.end]);
    }
    match removed_count {
        1 => {},
        0 => {
            return Err(Error::Bundle(format!(
                "Data identifier {data_identifier} is absent from PackageMetadata"
            )));
        },
        _ => {
            return Err(Error::Bundle(format!(
                "Data identifier {data_identifier} is duplicated in PackageMetadata"
            )));
        },
    }
    let decoded = protobuf::tsp::PackageMetadata::decode(output.as_slice())?;
    if decoded
        .datas
        .iter()
        .any(|data| data.identifier == data_identifier)
    {
        return Err(Error::InvalidFormat(
            "Removed DataInfo still decodes from PackageMetadata".to_owned(),
        ));
    }
    Ok(output)
}

fn patch_data_info(data: &[u8], digest: &[u8], materialized_length: u64) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(data)?;
    let mut output = Vec::with_capacity(data.len());
    let mut digest_count = 0usize;
    let mut length_count = 0usize;
    for field in fields {
        match field.number {
            2 => {
                if field.wire_type != 2 {
                    return Err(Error::InvalidFormat(
                        "DataInfo.digest has an invalid wire type".to_owned(),
                    ));
                }
                digest_count += 1;
                if digest_count > 1 {
                    return Err(Error::InvalidFormat(
                        "DataInfo contains duplicate digests".to_owned(),
                    ));
                }
                output.extend_from_slice(&data[field.start..field.key_end]);
                encode_varint_into(&mut output, digest.len() as u64);
                output.extend_from_slice(digest);
            },
            18 => {
                if field.wire_type != 0 {
                    return Err(Error::InvalidFormat(
                        "DataInfo.materialized_length has an invalid wire type".to_owned(),
                    ));
                }
                length_count += 1;
                if length_count > 1 {
                    return Err(Error::InvalidFormat(
                        "DataInfo contains duplicate materialized lengths".to_owned(),
                    ));
                }
                output.extend_from_slice(&data[field.start..field.key_end]);
                encode_varint_into(&mut output, materialized_length);
            },
            _ => output.extend_from_slice(&data[field.start..field.end]),
        }
    }
    if digest_count == 0 {
        encode_varint_into(&mut output, (2 << 3) | 2);
        encode_varint_into(&mut output, digest.len() as u64);
        output.extend_from_slice(digest);
    }
    if length_count == 0 {
        encode_varint_into(&mut output, 18 << 3);
        encode_varint_into(&mut output, materialized_length);
    }
    Ok(output)
}
