//! Media signatures, catalog validation, and iWork metadata wire edits.

use std::collections::{HashMap, HashSet, VecDeque};
use std::io::Write;
use std::path::{Component, Path};

use crate::package::IWorkPackage;
use crate::{Error, Result};
use litchi_iwa_common::varint::{decode_varint_from_bytes, encode_varint_into, encoded_len};
use litchi_iwa_common::wire::{WireField, parse_wire_fields_with_limits};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_index::ObjectId;
use litchi_iwa_protos::package_metadata_codec::{
    PackageMetadataVisitor, RewriteOptions, inspect_package_metadata_with_visitor,
};

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
        .map(|identifier| {
            ObjectId::try_from(identifier).map_err(|_| {
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
                        for data_identifier in data_metadata_identifiers(message.data.as_slice())? {
                            let data_identifier = MediaAssetId::try_from(data_identifier)?;
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

    let limits = metadata_wire_limits(metadata.datas.len())?;
    let mut assets = Vec::new();
    reserve_metadata(
        &mut assets,
        metadata.datas.len(),
        "media asset snapshots",
        limits,
    )?;
    let mut identifiers = std::collections::HashSet::new();
    identifiers
        .try_reserve(metadata.datas.len())
        .map_err(|_allocation| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "media asset identifiers",
                amount: metadata.datas.len(),
            })
        })?;
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

#[derive(Debug, Default)]
struct PackageMetadataSnapshot {
    components: Vec<ComponentSnapshot>,
    versioned_components: Vec<ComponentSnapshot>,
    datas: Vec<DataInfoSnapshot>,
    data_metadata_map: Option<u64>,
}

#[derive(Debug, Default)]
struct ComponentSnapshot {
    data_references: Vec<ComponentDataReferenceSnapshot>,
}

#[derive(Debug, Default)]
struct ComponentDataReferenceSnapshot {
    data_identifier: u64,
    object_reference_list: Vec<ObjectReferenceSnapshot>,
}

#[derive(Debug, Default)]
struct ObjectReferenceSnapshot {
    object_identifier: u64,
    count: u32,
}

#[derive(Debug, Default)]
struct DataInfoSnapshot {
    identifier: u64,
    digest: Vec<u8>,
    preferred_file_name: String,
    file_name: Option<String>,
    materialized_length: Option<u64>,
}

fn metadata_wire_limits(input_bytes: usize) -> Result<WireLimits> {
    let input_limit = input_bytes.clamp(1, WireLimits::MAX_INPUT_BYTES);
    WireLimits::default()
        .with_input_bytes(input_limit)
        .map_err(Into::into)
}

fn reserve_metadata<T>(
    values: &mut Vec<T>,
    additional: usize,
    resource: &'static str,
    limits: WireLimits,
) -> Result<()> {
    let requested = values.len().checked_add(additional).ok_or_else(|| {
        Error::InvalidFormat("Metadata collection size overflows usize".to_owned())
    })?;
    if requested > limits.max_fields() {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: requested,
            limit: limits.max_fields(),
        }));
    }
    values.try_reserve(additional).map_err(|_allocation| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: requested,
        })
    })
}

fn reserve_output(output: &mut Vec<u8>, additional: usize, limits: WireLimits) -> Result<()> {
    let requested = output
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat("Metadata output size overflows usize".to_owned()))?;
    if requested > limits.max_output_bytes() {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: requested,
            limit: limits.max_output_bytes(),
        }));
    }
    output.try_reserve(additional).map_err(|_allocation| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "media metadata output",
            amount: requested,
        })
    })
}

fn output_with_capacity(capacity: usize, limits: WireLimits) -> Result<Vec<u8>> {
    if capacity > limits.max_output_bytes() {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: capacity,
            limit: limits.max_output_bytes(),
        }));
    }
    let mut output = Vec::new();
    output.try_reserve_exact(capacity).map_err(|_allocation| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "media metadata output",
            amount: capacity,
        })
    })?;
    Ok(output)
}

fn extend_output(output: &mut Vec<u8>, bytes: &[u8], limits: WireLimits) -> Result<()> {
    reserve_output(output, bytes.len(), limits)?;
    output.extend_from_slice(bytes);
    Ok(())
}

struct NoopPackageMetadataVisitor;

impl PackageMetadataVisitor for NoopPackageMetadataVisitor {}

fn inspect_package_metadata_projection(metadata: &[u8]) -> Result<()> {
    // The private projection intentionally covers only the package envelope
    // and component registries. Keep its complete two-pass inspection bounded
    // by the source payload while leaving DataInfo and DataMetadataMap on the
    // source-preserving parser below.
    let source_limit = metadata.len().max(1);
    let options = RewriteOptions::new(
        source_limit,
        source_limit,
        source_limit.saturating_mul(2).max(1),
        metadata.len().saturating_mul(64).max(1),
        64,
        source_limit,
        source_limit,
        0,
    );
    inspect_package_metadata_with_visitor(metadata, options, &mut NoopPackageMetadataVisitor)
        .map(|_inspection| ())
        .map_err(|error| {
            Error::InvalidFormat(format!("PackageMetadata strict inspection failed: {error}"))
        })
}

fn strict_wire_fields(data: &[u8]) -> Result<Vec<WireField>> {
    // Parsing is intentionally structural here. Unknown records are copied
    // from their source spans by the rewrite helpers, so canonicalizing their
    // keys, length prefixes, or varint payloads would silently change bytes
    // that this codec does not own. Known fields validate their selected
    // framing/value at the point where they are decoded below.
    Ok(parse_wire_fields_with_limits(
        data,
        metadata_wire_limits(data.len())?,
    )?)
}

fn decode_package_metadata_payload(metadata: &[u8]) -> Result<PackageMetadataSnapshot> {
    inspect_package_metadata_projection(metadata)?;
    let fields = strict_wire_fields(metadata)?;
    let limits = metadata_wire_limits(metadata.len())?;
    let mut snapshot = PackageMetadataSnapshot::default();
    let mut last_object_identifier = None;
    let mut revision = false;
    let mut save_token = false;
    let mut preferred_package_type = false;
    for field in fields {
        match field.number() {
            1 => set_metadata_field(
                &mut last_object_identifier,
                field_varint(metadata, &field)?,
                "PackageMetadata.last_object_identifier",
            )?,
            2 => {
                let nested = field_payload(metadata, &field)?;
                let _ = strict_wire_fields(nested)?;
                set_metadata_seen(&mut revision, "PackageMetadata.revision")?;
            },
            3 => {
                reserve_metadata(
                    &mut snapshot.components,
                    1,
                    "PackageMetadata components",
                    limits,
                )?;
                snapshot
                    .components
                    .push(decode_component(field_payload(metadata, &field)?)?);
            },
            4 => {
                reserve_metadata(&mut snapshot.datas, 1, "PackageMetadata data infos", limits)?;
                snapshot
                    .datas
                    .push(decode_data_info(field_payload(metadata, &field)?)?);
            },
            10 => {
                if snapshot.data_metadata_map.is_some() {
                    return Err(Error::InvalidFormat(
                        "PackageMetadata contains duplicate data metadata map references"
                            .to_owned(),
                    ));
                }
                snapshot.data_metadata_map = Some(decode_reference_identifier(field_payload(
                    metadata, &field,
                )?)?);
            },
            5..=7 => validate_packed_varints(metadata, &field, u64::from(u32::MAX))?,
            8 => {
                let _ = field_varint(metadata, &field)?;
                set_metadata_seen(&mut save_token, "PackageMetadata.save_token")?;
            },
            9 => {
                let _ = field_varint(metadata, &field)?;
                set_metadata_seen(
                    &mut preferred_package_type,
                    "PackageMetadata.preferred_package_type",
                )?;
            },
            11 => {
                reserve_metadata(
                    &mut snapshot.versioned_components,
                    1,
                    "PackageMetadata versioned components",
                    limits,
                )?;
                snapshot
                    .versioned_components
                    .push(decode_component(field_payload(metadata, &field)?)?);
            },
            _ => {},
        }
    }
    let last_object_identifier = last_object_identifier.ok_or_else(|| {
        Error::InvalidFormat(
            "PackageMetadata is missing its required last object identifier".to_owned(),
        )
    })?;
    if last_object_identifier == 0 {
        return Err(Error::InvalidFormat(
            "PackageMetadata.last_object_identifier must be non-zero".to_owned(),
        ));
    }
    Ok(snapshot)
}

fn set_metadata_seen(slot: &mut bool, name: &str) -> Result<()> {
    if *slot {
        return Err(Error::InvalidFormat(format!("{name} is duplicated")));
    }
    *slot = true;
    Ok(())
}

fn validate_canonical_bool(data: &[u8], field: &WireField, name: &str) -> Result<()> {
    let value = field_varint(data, field)?;
    if value > 1 {
        return Err(Error::InvalidFormat(format!(
            "{name} is not a canonical bool"
        )));
    }
    Ok(())
}

fn validate_packed_varints(data: &[u8], field: &WireField, maximum: u64) -> Result<()> {
    let mut payload = field_payload(data, field)?;
    while !payload.is_empty() {
        let (value, length) = decode_varint_from_bytes(payload).map_err(|error| {
            Error::InvalidFormat(format!(
                "Protobuf packed field {} contains an invalid varint: {error}",
                field.number()
            ))
        })?;
        if length != encoded_len(value) {
            return Err(Error::InvalidFormat(format!(
                "Protobuf packed field {} contains a noncanonical varint",
                field.number()
            )));
        }
        if value > maximum {
            return Err(Error::InvalidFormat(format!(
                "Protobuf packed field {} value exceeds its schema width",
                field.number()
            )));
        }
        payload = &payload[length..];
    }
    Ok(())
}

fn decode_component(data: &[u8]) -> Result<ComponentSnapshot> {
    let fields = strict_wire_fields(data)?;
    let mut component = ComponentSnapshot::default();
    let mut identifier = None;
    let mut preferred_locator = false;
    let mut locator = false;
    let mut is_stored_outside_object_archive = false;
    let mut save_token = false;
    let mut compression_algorithm = false;
    let mut can_be_dropped = false;
    let mut is_wasteful = false;
    let mut required_package_identifier = false;
    for field in fields {
        match field.number() {
            1 => set_metadata_field(
                &mut identifier,
                field_varint(data, &field)?,
                "ComponentInfo.identifier",
            )?,
            2 => {
                let _ = metadata_utf8(data, &field, "ComponentInfo.preferred_locator")?;
                set_metadata_seen(&mut preferred_locator, "ComponentInfo.preferred_locator")?;
            },
            3 => {
                let _ = metadata_utf8(data, &field, "ComponentInfo.locator")?;
                set_metadata_seen(&mut locator, "ComponentInfo.locator")?;
            },
            4 | 5 | 14 | 15 => validate_packed_varints(data, &field, u64::from(u32::MAX))?,
            6 | 13 | 18 => {
                let _ = field_payload(data, &field)?;
            },
            7 => {
                reserve_metadata(
                    &mut component.data_references,
                    1,
                    "ComponentInfo data references",
                    metadata_wire_limits(data.len())?,
                )?;
                component
                    .data_references
                    .push(decode_component_data_reference(field_payload(
                        data, &field,
                    )?)?);
            },
            10 => {
                validate_canonical_bool(
                    data,
                    &field,
                    "ComponentInfo.is_stored_outside_object_archive",
                )?;
                set_metadata_seen(
                    &mut is_stored_outside_object_archive,
                    "ComponentInfo.is_stored_outside_object_archive",
                )?;
            },
            11 => {
                let _ = field_payload(data, &field)?;
            },
            12 => {
                let _ = field_varint(data, &field)?;
                set_metadata_seen(&mut save_token, "ComponentInfo.save_token")?;
            },
            16 => {
                let value = field_varint(data, &field)?;
                u32::try_from(value).map_err(|_error| {
                    Error::InvalidFormat(
                        "ComponentInfo.compression_algorithm exceeds u32".to_owned(),
                    )
                })?;
                set_metadata_seen(
                    &mut compression_algorithm,
                    "ComponentInfo.compression_algorithm",
                )?;
            },
            17 => {
                validate_canonical_bool(data, &field, "ComponentInfo.can_be_dropped")?;
                set_metadata_seen(&mut can_be_dropped, "ComponentInfo.can_be_dropped")?;
            },
            19 => {
                validate_canonical_bool(data, &field, "ComponentInfo.is_wasteful")?;
                set_metadata_seen(&mut is_wasteful, "ComponentInfo.is_wasteful")?;
            },
            20 => validate_packed_varints(data, &field, u64::MAX)?,
            21 => {
                let _ = field_varint(data, &field)?;
                set_metadata_seen(
                    &mut required_package_identifier,
                    "ComponentInfo.required_package_identifier",
                )?;
            },
            _ => {},
        }
    }
    let _identifier = identifier.ok_or_else(|| {
        Error::InvalidFormat("ComponentInfo is missing its required identifier".to_owned())
    })?;
    if !preferred_locator {
        return Err(Error::InvalidFormat(
            "ComponentInfo is missing its required preferred locator".to_owned(),
        ));
    }
    Ok(component)
}

fn decode_component_data_reference(data: &[u8]) -> Result<ComponentDataReferenceSnapshot> {
    let fields = strict_wire_fields(data)?;
    let limits = metadata_wire_limits(data.len())?;
    let mut data_identifier = None;
    let mut object_reference_list = Vec::new();
    for field in fields {
        match field.number() {
            1 => set_metadata_field(
                &mut data_identifier,
                field_varint(data, &field)?,
                "ComponentDataReference.data_identifier",
            )?,
            2 => {
                reserve_metadata(
                    &mut object_reference_list,
                    1,
                    "ComponentDataReference object references",
                    limits,
                )?;
                object_reference_list.push(decode_object_reference(field_payload(data, &field)?)?);
            },
            _ => {},
        }
    }
    Ok(ComponentDataReferenceSnapshot {
        data_identifier: data_identifier.ok_or_else(|| {
            Error::InvalidFormat(
                "ComponentDataReference is missing its required data identifier".to_owned(),
            )
        })?,
        object_reference_list,
    })
}

fn decode_object_reference(data: &[u8]) -> Result<ObjectReferenceSnapshot> {
    let fields = strict_wire_fields(data)?;
    let mut object_identifier = None;
    let mut count = None;
    for field in fields {
        match field.number() {
            1 => set_metadata_field(
                &mut object_identifier,
                field_varint(data, &field)?,
                "ComponentDataReference.ObjectReference.object_identifier",
            )?,
            2 => {
                let value = field_varint(data, &field)?;
                let value = u32::try_from(value).map_err(|_error| {
                    Error::InvalidFormat(
                        "ComponentDataReference.ObjectReference.count exceeds u32".to_owned(),
                    )
                })?;
                set_metadata_field(
                    &mut count,
                    value,
                    "ComponentDataReference.ObjectReference.count",
                )?;
            },
            _ => {},
        }
    }
    Ok(ObjectReferenceSnapshot {
        object_identifier: object_identifier.ok_or_else(|| {
            Error::InvalidFormat(
                "ComponentDataReference.ObjectReference is missing its required object identifier"
                    .to_owned(),
            )
        })?,
        count: count.ok_or_else(|| {
            Error::InvalidFormat(
                "ComponentDataReference.ObjectReference is missing its required count".to_owned(),
            )
        })?,
    })
}

fn decode_data_info(data: &[u8]) -> Result<DataInfoSnapshot> {
    let fields = strict_wire_fields(data)?;
    let limits = metadata_wire_limits(data.len())?;
    let mut identifier = None;
    let mut digest = None;
    let mut preferred_file_name = None;
    let mut file_name = None;
    let mut materialized_length = None;
    let mut document_resource_locator = false;
    let mut source_bookmark_data = false;
    let mut remote_url = false;
    let mut can_download = false;
    let mut download_priority = false;
    let mut attributes = false;
    let mut encryption_info = false;
    let mut last_mismatched_digest = false;
    let mut unmaterialized_ranges = false;
    let mut remote_data_length = false;
    let mut remote_data_has_package_storage = false;
    let mut upload_status = false;
    let mut remote_data_mtime = false;
    let mut pasteboard_external_file_path = false;
    for field in fields {
        match field.number() {
            1 => set_metadata_field(
                &mut identifier,
                field_varint(data, &field)?,
                "DataInfo.identifier",
            )?,
            2 => set_metadata_field(
                &mut digest,
                metadata_bytes(data, &field, "DataInfo.digest", limits)?,
                "DataInfo.digest",
            )?,
            3 => set_metadata_field(
                &mut preferred_file_name,
                metadata_string(data, &field, "DataInfo.preferred_file_name")?,
                "DataInfo.preferred_file_name",
            )?,
            4 => set_metadata_field(
                &mut file_name,
                metadata_string(data, &field, "DataInfo.file_name")?,
                "DataInfo.file_name",
            )?,
            5 | 7 | 99 => {
                let name = match field.number() {
                    5 => "DataInfo.document_resource_locator",
                    7 => "DataInfo.remote_url",
                    _ => "DataInfo.pasteboard_external_file_path",
                };
                let _ = metadata_utf8(data, &field, name)?;
                set_metadata_seen(
                    match field.number() {
                        5 => &mut document_resource_locator,
                        7 => &mut remote_url,
                        _ => &mut pasteboard_external_file_path,
                    },
                    name,
                )?;
            },
            6 | 12 => {
                let name = if field.number() == 6 {
                    "DataInfo.source_bookmark_data"
                } else {
                    "DataInfo.last_mismatched_digest"
                };
                let _ = field_payload(data, &field)?;
                set_metadata_seen(
                    if field.number() == 6 {
                        &mut source_bookmark_data
                    } else {
                        &mut last_mismatched_digest
                    },
                    name,
                )?;
            },
            8 | 15 => {
                let name = if field.number() == 8 {
                    "DataInfo.can_download"
                } else {
                    "DataInfo.remote_data_has_package_storage"
                };
                validate_canonical_bool(data, &field, name)?;
                set_metadata_seen(
                    if field.number() == 8 {
                        &mut can_download
                    } else {
                        &mut remote_data_has_package_storage
                    },
                    name,
                )?;
            },
            9 | 16 => {
                let name = if field.number() == 9 {
                    "DataInfo.download_priority"
                } else {
                    "DataInfo.upload_status"
                };
                let _ = field_varint(data, &field)?;
                set_metadata_seen(
                    if field.number() == 9 {
                        &mut download_priority
                    } else {
                        &mut upload_status
                    },
                    name,
                )?;
            },
            10 | 11 | 13 => {
                let name = match field.number() {
                    10 => "DataInfo.attributes",
                    11 => "DataInfo.encryption_info",
                    _ => "DataInfo.unmaterialized_ranges",
                };
                let nested = field_payload(data, &field)?;
                let _ = strict_wire_fields(nested)?;
                set_metadata_seen(
                    match field.number() {
                        10 => &mut attributes,
                        11 => &mut encryption_info,
                        _ => &mut unmaterialized_ranges,
                    },
                    name,
                )?;
            },
            14 => {
                let _ = field_varint(data, &field)?;
                set_metadata_seen(&mut remote_data_length, "DataInfo.remote_data_length")?;
            },
            17 => {
                if field.wire_type() != 1 {
                    return Err(Error::InvalidFormat(format!(
                        "DataInfo.remote_data_mtime has an invalid wire type: {}",
                        field.wire_type()
                    )));
                }
                field.validate_canonical_key(data)?;
                if field.checked_payload(data)?.len() != 8 {
                    return Err(Error::InvalidFormat(
                        "DataInfo.remote_data_mtime is not a fixed64 value".to_owned(),
                    ));
                }
                set_metadata_seen(&mut remote_data_mtime, "DataInfo.remote_data_mtime")?;
            },
            18 => set_metadata_field(
                &mut materialized_length,
                field_varint(data, &field)?,
                "DataInfo.materialized_length",
            )?,
            _ => {},
        }
    }
    Ok(DataInfoSnapshot {
        identifier: identifier.ok_or_else(|| {
            Error::InvalidFormat("DataInfo is missing its required identifier".to_owned())
        })?,
        digest: digest.ok_or_else(|| {
            Error::InvalidFormat("DataInfo is missing its required digest".to_owned())
        })?,
        preferred_file_name: preferred_file_name.ok_or_else(|| {
            Error::InvalidFormat("DataInfo is missing its required preferred file name".to_owned())
        })?,
        file_name,
        materialized_length,
    })
}

fn decode_reference_identifier(data: &[u8]) -> Result<u64> {
    let fields = strict_wire_fields(data)?;
    let mut identifier = None;
    let mut deprecated_type = false;
    let mut deprecated_is_external = false;
    for field in fields {
        match field.number() {
            1 => set_metadata_field(
                &mut identifier,
                field_varint(data, &field)?,
                "Reference.identifier",
            )?,
            2 => {
                let _ = field_varint(data, &field)?;
                set_metadata_seen(&mut deprecated_type, "Reference.deprecated_type")?;
            },
            3 => {
                validate_canonical_bool(data, &field, "Reference.deprecated_is_external")?;
                set_metadata_seen(
                    &mut deprecated_is_external,
                    "Reference.deprecated_is_external",
                )?;
            },
            _ => {},
        }
    }
    identifier.ok_or_else(|| {
        Error::InvalidFormat("Reference is missing its required identifier".to_owned())
    })
}

fn data_metadata_identifiers(data: &[u8]) -> Result<Vec<u64>> {
    let fields = strict_wire_fields(data)?;
    let limits = metadata_wire_limits(data.len())?;
    let mut identifiers = Vec::new();
    for field in fields {
        if field.number() == 1 {
            reserve_metadata(&mut identifiers, 1, "DataMetadataMap entries", limits)?;
            identifiers.push(decode_data_metadata_entry(field_payload(data, &field)?)?);
        }
    }
    Ok(identifiers)
}

fn decode_data_metadata_entry(data: &[u8]) -> Result<u64> {
    let fields = strict_wire_fields(data)?;
    let mut data_identifier = None;
    let mut data_metadata = None;
    for field in fields {
        match field.number() {
            1 => set_metadata_field(
                &mut data_identifier,
                field_varint(data, &field)?,
                "DataMetadataMap.DataMetadataMapEntry.data_identifier",
            )?,
            2 => set_metadata_field(
                &mut data_metadata,
                decode_reference_identifier(field_payload(data, &field)?)?,
                "DataMetadataMap.DataMetadataMapEntry.data_metadata",
            )?,
            _ => {},
        }
    }
    let _data_metadata = data_metadata.ok_or_else(|| {
        Error::InvalidFormat(
            "DataMetadataMap entry is missing its required data metadata reference".to_owned(),
        )
    })?;
    data_identifier.ok_or_else(|| {
        Error::InvalidFormat(
            "DataMetadataMap entry is missing its required data identifier".to_owned(),
        )
    })
}

fn set_metadata_field<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::InvalidFormat(format!("{name} is duplicated")));
    }
    Ok(())
}

fn metadata_string(data: &[u8], field: &WireField, name: &str) -> Result<String> {
    let value = metadata_utf8(data, field, name)?;
    let mut output = String::new();
    output.try_reserve(value.len()).map_err(|_allocation| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "media metadata string",
            amount: value.len(),
        })
    })?;
    output.push_str(value);
    Ok(output)
}

fn metadata_bytes(
    data: &[u8],
    field: &WireField,
    _name: &str,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let value = field_payload(data, field)?;
    let mut output = Vec::new();
    if value.len() > limits.max_input_bytes() {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed: value.len(),
            limit: limits.max_input_bytes(),
        }));
    }
    output
        .try_reserve_exact(value.len())
        .map_err(|_allocation| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "media metadata bytes",
                amount: value.len(),
            })
        })?;
    output.extend_from_slice(value);
    Ok(output)
}

fn metadata_utf8<'a>(data: &'a [u8], field: &WireField, name: &str) -> Result<&'a str> {
    let payload = field_payload(data, field)?;
    std::str::from_utf8(payload)
        .map_err(|_error| Error::InvalidFormat(format!("{name} is not valid UTF-8")))
}

fn decode_package_metadata(package: &IWorkPackage) -> Result<PackageMetadataSnapshot> {
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
    decode_package_metadata_payload(
        payload.ok_or_else(|| Error::Bundle("PackageMetadata payload was not found".to_owned()))?,
    )
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

pub(crate) fn field_payload<'a>(data: &'a [u8], field: &WireField) -> Result<&'a [u8]> {
    if field.wire_type() != 2 {
        return Err(Error::InvalidFormat(format!(
            "Protobuf field {} is not length-delimited",
            field.number()
        )));
    }
    field.validate_canonical_framing(data)?;
    field.checked_payload(data).map_err(Error::from)
}

fn field_varint(data: &[u8], field: &WireField) -> Result<u64> {
    if field.wire_type() != 0 {
        return Err(Error::InvalidFormat(format!(
            "Protobuf field {} is not a varint",
            field.number()
        )));
    }
    field.validate_canonical_key(data)?;
    let payload = field.checked_payload(data)?;
    let (value, length) = decode_varint_from_bytes(payload)
        .map_err(|error| Error::InvalidFormat(format!("Invalid protobuf varint: {error}")))?;
    if length != payload.len() || length != encoded_len(value) {
        return Err(Error::InvalidFormat(format!(
            "Protobuf field {} contains a noncanonical varint",
            field.number()
        )));
    }
    Ok(value)
}

fn data_info_identifier(data: &[u8]) -> Result<u64> {
    let fields = strict_wire_fields(data)?;
    let mut identifier = None;
    for field in fields.iter().filter(|field| field.number() == 1) {
        set_metadata_field(
            &mut identifier,
            field_varint(data, field)?,
            "DataInfo.identifier",
        )?;
    }
    identifier.ok_or_else(|| {
        Error::InvalidFormat("DataInfo is missing its required identifier".to_owned())
    })
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
    let _source = decode_package_metadata_payload(metadata)?;
    let fields = strict_wire_fields(metadata)?;
    let limits = metadata_wire_limits(metadata.len())?;
    let mut output = output_with_capacity(metadata.len(), limits)?;
    let mut patched_count = 0usize;
    for field in fields {
        if field.number() == 4 {
            if field.wire_type() != 2 {
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
                extend_output(
                    &mut output,
                    &metadata[field.start()..field.key_end()],
                    limits,
                )?;
                append_varint_to_output(
                    &mut output,
                    u64::try_from(patched.len()).map_err(|_error| {
                        Error::InvalidFormat("Patched DataInfo length exceeds u64".to_owned())
                    })?,
                    limits,
                )?;
                extend_output(&mut output, &patched, limits)?;
                continue;
            }
        }
        extend_output(&mut output, &metadata[field.start()..field.end()], limits)?;
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
    let decoded = decode_package_metadata_payload(output.as_slice())?;
    let mut matches = decoded
        .datas
        .iter()
        .filter(|data| data.identifier == data_identifier);
    let Some(matched) = matches.next() else {
        return Err(Error::InvalidFormat(
            "Patched PackageMetadata did not decode to the requested values".to_owned(),
        ));
    };
    if matches.next().is_some()
        || matched.digest != digest
        || matched.materialized_length != Some(materialized_length)
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
    let decoded = decode_package_metadata_payload(metadata)?;
    if decoded
        .datas
        .iter()
        .any(|data| data.identifier == data_identifier)
    {
        return Err(Error::Bundle(format!(
            "Data identifier {data_identifier} already exists"
        )));
    }

    let limits = metadata_wire_limits(metadata.len())?;
    let mut data_info = output_with_capacity(0, limits)?;
    append_wire_varint_limited(&mut data_info, 1, data_identifier, limits)?;
    append_wire_bytes_limited(&mut data_info, 2, digest, limits)?;
    append_wire_bytes_limited(&mut data_info, 3, preferred_filename.as_bytes(), limits)?;
    append_wire_bytes_limited(&mut data_info, 4, file_name.as_bytes(), limits)?;
    append_wire_varint_limited(&mut data_info, 18, materialized_length, limits)?;

    // Appending a repeated field is protobuf-canonical and avoids rewriting any
    // pre-existing metadata field, including unknown extensions.
    let mut output = output_with_capacity(metadata.len(), limits)?;
    extend_output(&mut output, metadata, limits)?;
    append_wire_bytes_limited(&mut output, 4, &data_info, limits)?;
    let verified = decode_package_metadata_payload(output.as_slice())?;
    let mut remaining = verified
        .datas
        .iter()
        .filter(|data| data.identifier == data_identifier);
    let Some(inserted) = remaining.next() else {
        return Err(Error::InvalidFormat(
            "Appended DataInfo did not decode to the requested values".to_owned(),
        ));
    };
    if remaining.next().is_some()
        || inserted.digest != digest
        || inserted.preferred_file_name != preferred_filename
        || inserted.file_name.as_deref() != Some(file_name)
        || inserted.materialized_length != Some(materialized_length)
    {
        return Err(Error::InvalidFormat(
            "Appended DataInfo did not decode to the requested values".to_owned(),
        ));
    }
    Ok(output)
}

#[cfg(test)]
fn append_wire_varint(output: &mut Vec<u8>, field_number: u64, value: u64) {
    encode_varint_into(output, field_number << 3);
    encode_varint_into(output, value);
}

#[cfg(test)]
fn append_wire_bytes(output: &mut Vec<u8>, field_number: u64, value: &[u8]) {
    encode_varint_into(output, (field_number << 3) | 2);
    encode_varint_into(output, value.len() as u64);
    output.extend_from_slice(value);
}

fn append_varint_to_output(output: &mut Vec<u8>, value: u64, limits: WireLimits) -> Result<()> {
    reserve_output(output, encoded_len(value), limits)?;
    encode_varint_into(output, value);
    Ok(())
}

fn append_wire_varint_limited(
    output: &mut Vec<u8>,
    field_number: u64,
    value: u64,
    limits: WireLimits,
) -> Result<()> {
    let key = field_number
        .checked_shl(3)
        .ok_or_else(|| Error::InvalidFormat("Metadata field number overflows u64".to_owned()))?;
    reserve_output(
        output,
        encoded_len(key)
            .checked_add(encoded_len(value))
            .ok_or_else(|| {
                Error::InvalidFormat("Metadata varint size overflows usize".to_owned())
            })?,
        limits,
    )?;
    encode_varint_into(output, key);
    encode_varint_into(output, value);
    Ok(())
}

fn append_wire_bytes_limited(
    output: &mut Vec<u8>,
    field_number: u64,
    value: &[u8],
    limits: WireLimits,
) -> Result<()> {
    let key = field_number
        .checked_shl(3)
        .and_then(|key| key.checked_add(2))
        .ok_or_else(|| Error::InvalidFormat("Metadata field number overflows u64".to_owned()))?;
    let length = u64::try_from(value.len())
        .map_err(|_error| Error::InvalidFormat("Metadata field length exceeds u64".to_owned()))?;
    let additional = encoded_len(key)
        .checked_add(encoded_len(length))
        .and_then(|size| size.checked_add(value.len()))
        .ok_or_else(|| Error::InvalidFormat("Metadata bytes size overflows usize".to_owned()))?;
    reserve_output(output, additional, limits)?;
    encode_varint_into(output, key);
    encode_varint_into(output, length);
    output.extend_from_slice(value);
    Ok(())
}

pub(crate) fn remove_data_info(metadata: &[u8], data_identifier: u64) -> Result<Vec<u8>> {
    let _source = decode_package_metadata_payload(metadata)?;
    let fields = strict_wire_fields(metadata)?;
    let limits = metadata_wire_limits(metadata.len())?;
    let mut output = output_with_capacity(metadata.len(), limits)?;
    let mut removed_count = 0usize;
    for field in fields {
        if field.number() == 4 {
            if field.wire_type() != 2 {
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
        extend_output(&mut output, &metadata[field.start()..field.end()], limits)?;
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
    let decoded = decode_package_metadata_payload(output.as_slice())?;
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
    let _source = decode_data_info(data)?;
    let fields = strict_wire_fields(data)?;
    let limits = metadata_wire_limits(data.len())?;
    let mut output = output_with_capacity(data.len(), limits)?;
    let mut digest_count = 0usize;
    let mut length_count = 0usize;
    for field in fields {
        match field.number() {
            2 => {
                if field.wire_type() != 2 {
                    return Err(Error::InvalidFormat(
                        "DataInfo.digest has an invalid wire type".to_owned(),
                    ));
                }
                let _ = field_payload(data, &field)?;
                digest_count += 1;
                if digest_count > 1 {
                    return Err(Error::InvalidFormat(
                        "DataInfo contains duplicate digests".to_owned(),
                    ));
                }
                extend_output(&mut output, &data[field.start()..field.key_end()], limits)?;
                append_varint_to_output(
                    &mut output,
                    u64::try_from(digest.len()).map_err(|_error| {
                        Error::InvalidFormat("Digest length exceeds u64".to_owned())
                    })?,
                    limits,
                )?;
                extend_output(&mut output, digest, limits)?;
            },
            18 => {
                if field.wire_type() != 0 {
                    return Err(Error::InvalidFormat(
                        "DataInfo.materialized_length has an invalid wire type".to_owned(),
                    ));
                }
                let _ = field_varint(data, &field)?;
                length_count += 1;
                if length_count > 1 {
                    return Err(Error::InvalidFormat(
                        "DataInfo contains duplicate materialized lengths".to_owned(),
                    ));
                }
                extend_output(&mut output, &data[field.start()..field.key_end()], limits)?;
                append_varint_to_output(&mut output, materialized_length, limits)?;
            },
            _ => extend_output(&mut output, &data[field.start()..field.end()], limits)?,
        }
    }
    if digest_count == 0 {
        append_wire_bytes_limited(&mut output, 2, digest, limits)?;
    }
    if length_count == 0 {
        append_wire_varint_limited(&mut output, 18, materialized_length, limits)?;
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn data_info(identifier: u64) -> Vec<u8> {
        let mut data_info = Vec::new();
        append_wire_varint(&mut data_info, 1, identifier);
        append_wire_bytes(&mut data_info, 2, &[0x11; 20]);
        append_wire_bytes(&mut data_info, 3, b"image.png");
        append_wire_bytes(&mut data_info, 4, b"image-7.png");
        append_wire_varint(&mut data_info, 18, 7);
        data_info
    }

    fn package_metadata_with_data_info(data_info: &[u8]) -> Vec<u8> {
        let mut metadata = Vec::new();
        append_wire_varint(&mut metadata, 1, 100);
        append_wire_bytes(&mut metadata, 4, data_info);
        metadata
    }

    fn uuid_entry(identifier: u64, lower: u64, upper: u64) -> Vec<u8> {
        let mut uuid = Vec::new();
        append_wire_varint(&mut uuid, 1, lower);
        append_wire_varint(&mut uuid, 2, upper);

        let mut entry = Vec::new();
        append_wire_varint(&mut entry, 1, identifier);
        append_wire_bytes(&mut entry, 2, &uuid);
        entry
    }

    #[test]
    fn legacy_uuid_registry_metadata_fits_strict_inspection_budget() {
        let mut component = Vec::new();
        append_wire_varint(&mut component, 1, 4);
        append_wire_bytes(&mut component, 2, b"Slide-4");
        for identifier in [4, 5, 6, 70, 71, 72, 73] {
            let entry = uuid_entry(identifier, identifier, identifier + 100);
            append_wire_bytes(&mut component, 11, &entry);
        }

        let mut metadata = Vec::new();
        append_wire_varint(&mut metadata, 1, 100);
        append_wire_bytes(&mut metadata, 3, &component);

        let snapshot = decode_package_metadata_payload(&metadata).unwrap();
        assert_eq!(snapshot.components.len(), 1);
        assert!(snapshot.datas.is_empty());
    }

    #[test]
    fn rewrites_preserve_noncanonical_unknown_varints() {
        let mut data_info = data_info(7);
        // Field 50 is unknown to the media codec. Its value intentionally
        // uses an overlong varint representation and must survive a patch.
        data_info.extend_from_slice(&[0x90, 0x03, 0x81, 0x00]);
        let metadata = package_metadata_with_data_info(&data_info);
        let patched = patch_package_metadata(&metadata, 7, &[0x22; 20], 11).unwrap();
        assert!(
            patched
                .windows(4)
                .any(|window| window == [0x90, 0x03, 0x81, 0x00])
        );

        let removed = remove_data_info(&patched, 7).unwrap();
        assert!(
            !removed
                .windows(4)
                .any(|window| window == [0x90, 0x03, 0x81, 0x00])
        );
    }

    #[test]
    fn known_wrong_wire_and_duplicate_fields_are_rejected() {
        let mut invalid_data_info = data_info(7);
        // DataInfo.remote_data_length is a known uint64 field, not bytes.
        append_wire_bytes(&mut invalid_data_info, 14, b"wrong-wire");
        let metadata = package_metadata_with_data_info(&invalid_data_info);
        assert!(decode_package_metadata_payload(&metadata).is_err());

        let mut duplicate_reference = Vec::new();
        append_wire_varint(&mut duplicate_reference, 1, 9);
        append_wire_varint(&mut duplicate_reference, 1, 10);
        assert!(decode_reference_identifier(&duplicate_reference).is_err());
    }
}
