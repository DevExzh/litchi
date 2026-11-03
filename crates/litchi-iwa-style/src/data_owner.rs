//! The package-metadata data-owner transitions of embedded media (record
//! 0422): a `TSP.DataInfo` per `Data/` member and, per component, the
//! objects that own a data reference (the host's
//! `data_reference_registry` and `append_data_info`). Lifted from Pages
//! (records 0371, 0413b, 0415) so the three format families share them.
//!
//! Payload functions edit a `TSP.PackageMetadata` payload; [`own`] and
//! [`release`] apply one owner transition to the registry of a
//! [`Workspace`].

use litchi_iwa_common::Result;
use litchi_iwa_common::wire::{
    WireView, patch_varint_field, remove_repeated_length_delimited_field_where,
    repeated_length_delimited_payloads, transform_length_delimited_fields_at_path,
};

use crate::graph::Workspace;

const METADATA_COMPONENTS_FIELD: u32 = 3;
const COMPONENT_IDENTIFIER_FIELD: u32 = 1;

fn invalid(what: &str) -> litchi_iwa_common::Error {
    litchi_iwa_common::Error::InvalidFormat(format!("{what} is invalid"))
}

fn varint_of(view: &WireView<'_>, number: u32) -> Option<u64> {
    view.fields()
        .find(|field| field.number() == number)
        .and_then(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
        .map(|(value, _)| value)
}

const METADATA_VERSIONED_COMPONENTS_FIELD: u32 = 11;
const METADATA_DATAS_FIELD: u32 = 4;
const COMPONENT_DATA_REFERENCES_FIELD: u32 = 7;
const DATA_REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const DATA_REFERENCE_OWNERS_FIELD: u32 = 2;
const OWNER_OBJECT_FIELD: u32 = 1;
const OWNER_COUNT_FIELD: u32 = 2;
const DATA_INFO_IDENTIFIER_FIELD: u32 = 1;
const DATA_INFO_FILE_NAME_FIELD: u32 = 4;

fn view_varint(payload: &[u8], number: u32) -> Result<Option<u64>> {
    let view = WireView::parse(payload, litchi_iwa_common::WireLimits::default())?;
    Ok(varint_of(&view, number))
}

/// Add one count to, or remove one count from, `component`'s
/// data-reference owner (`data`, `object`) in a metadata payload (the host's
/// `{add,remove}_component_data_reference`): a removal that empties the
/// only owner drops the whole data reference, one that empties an owner
/// drops the owner, and an addition appends a new reference or owner when
/// none exists.
pub fn adjust_data_reference(
    payload: &[u8],
    component: u64,
    data: u64,
    object: u64,
    add: bool,
) -> Result<Vec<u8>> {
    let mut matched = 0usize;
    let mut output = payload.to_vec();
    for field in [
        METADATA_COMPONENTS_FIELD,
        METADATA_VERSIONED_COMPONENTS_FIELD,
    ] {
        output = transform_length_delimited_fields_at_path(&output, &[field], |info| {
            if view_varint(info, COMPONENT_IDENTIFIER_FIELD)? != Some(component) {
                return Ok(info.to_vec());
            }
            matched += 1;
            adjust_component(info, data, object, add)
        })?;
    }
    if matched != 1 {
        return Err(invalid("a data reference"));
    }
    Ok(output)
}

fn adjust_component(info: &[u8], data: u64, object: u64, add: bool) -> Result<Vec<u8>> {
    let invalid = || litchi_iwa_common::Error::InvalidFormat("data reference".to_owned());
    let references = repeated_length_delimited_payloads(info, COMPONENT_DATA_REFERENCES_FIELD)?;
    let mut matches = Vec::new();
    for reference in references {
        if view_varint(reference, DATA_REFERENCE_IDENTIFIER_FIELD)? == Some(data) {
            matches.push(reference);
        }
    }
    let is_data = |reference: &[u8]| -> Result<bool> {
        Ok(view_varint(reference, DATA_REFERENCE_IDENTIFIER_FIELD)? == Some(data))
    };
    match (add, matches.as_slice()) {
        (true, []) => {
            let mut owner = Vec::new();
            litchi_iwa_common::wire::append_varint_field(&mut owner, OWNER_OBJECT_FIELD, object)?;
            litchi_iwa_common::wire::append_varint_field(&mut owner, OWNER_COUNT_FIELD, 1)?;
            let mut reference = Vec::new();
            litchi_iwa_common::wire::append_varint_field(
                &mut reference,
                DATA_REFERENCE_IDENTIFIER_FIELD,
                data,
            )?;
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut reference,
                DATA_REFERENCE_OWNERS_FIELD,
                &owner,
                litchi_iwa_common::WireLimits::default(),
            )?;
            litchi_iwa_common::wire::append_repeated_length_delimited_field(
                info,
                COMPONENT_DATA_REFERENCES_FIELD,
                &reference,
            )
        },
        (_, [reference]) => {
            let owners =
                repeated_length_delimited_payloads(reference, DATA_REFERENCE_OWNERS_FIELD)?;
            let mut found = Vec::new();
            for owner in &owners {
                if view_varint(owner, OWNER_OBJECT_FIELD)? == Some(object) {
                    found.push(view_varint(owner, OWNER_COUNT_FIELD)?.unwrap_or(0));
                }
            }
            if !add && found == [1] && owners.len() == 1 {
                return remove_repeated_length_delimited_field_where(
                    info,
                    COMPONENT_DATA_REFERENCES_FIELD,
                    is_data,
                );
            }
            transform_length_delimited_fields_at_path(
                info,
                &[COMPONENT_DATA_REFERENCES_FIELD],
                |reference| {
                    if !is_data(reference)? {
                        return Ok(reference.to_vec());
                    }
                    adjust_owner(reference, object, add, &found)
                },
            )
        },
        _ => Err(invalid()),
    }
}

fn adjust_owner(reference: &[u8], object: u64, add: bool, found: &[u64]) -> Result<Vec<u8>> {
    let is_owner = |owner: &[u8]| -> Result<bool> {
        Ok(view_varint(owner, OWNER_OBJECT_FIELD)? == Some(object))
    };
    match (add, found) {
        (true, []) => {
            let mut owner = Vec::new();
            litchi_iwa_common::wire::append_varint_field(&mut owner, OWNER_OBJECT_FIELD, object)?;
            litchi_iwa_common::wire::append_varint_field(&mut owner, OWNER_COUNT_FIELD, 1)?;
            litchi_iwa_common::wire::append_repeated_length_delimited_field(
                reference,
                DATA_REFERENCE_OWNERS_FIELD,
                &owner,
            )
        },
        (false, [1]) => remove_repeated_length_delimited_field_where(
            reference,
            DATA_REFERENCE_OWNERS_FIELD,
            is_owner,
        ),
        (_, [count]) => {
            let count = if add {
                count.checked_add(1)
            } else {
                count.checked_sub(1)
            }
            .ok_or_else(|| {
                litchi_iwa_common::Error::InvalidFormat("data reference count".to_owned())
            })?;
            transform_length_delimited_fields_at_path(
                reference,
                &[DATA_REFERENCE_OWNERS_FIELD],
                |owner| {
                    if !is_owner(owner)? {
                        return Ok(owner.to_vec());
                    }
                    patch_varint_field(owner, OWNER_COUNT_FIELD, true, Some(count))
                },
            )
        },
        _ => Err(litchi_iwa_common::Error::InvalidFormat(
            "data reference owner".to_owned(),
        )),
    }
}

/// The data identifiers any component (current or versioned) still
/// references in a metadata payload.
pub fn referenced_data(payload: &[u8]) -> Result<std::collections::HashSet<u64>> {
    let mut referenced = std::collections::HashSet::new();
    for field in [
        METADATA_COMPONENTS_FIELD,
        METADATA_VERSIONED_COMPONENTS_FIELD,
    ] {
        for info in repeated_length_delimited_payloads(payload, field)? {
            for reference in
                repeated_length_delimited_payloads(info, COMPONENT_DATA_REFERENCES_FIELD)?
            {
                if let Some(data) = view_varint(reference, DATA_REFERENCE_IDENTIFIER_FIELD)? {
                    referenced.insert(data);
                }
            }
        }
    }
    Ok(referenced)
}

/// The `Data/` member name of a data entry, from its `DataInfo.file_name`.
pub fn data_member(payload: &[u8], data: u64) -> Result<Option<String>> {
    let mut found = None;
    for info in repeated_length_delimited_payloads(payload, METADATA_DATAS_FIELD)? {
        if view_varint(info, DATA_INFO_IDENTIFIER_FIELD)? != Some(data) {
            continue;
        }
        if found.is_some() {
            return Err(invalid("a data reference"));
        }
        let view = WireView::parse(info, litchi_iwa_common::WireLimits::default())?;
        let name = view
            .fields()
            .find(|field| field.number() == DATA_INFO_FILE_NAME_FIELD)
            .map(|field| std::str::from_utf8(field.payload()).map(str::to_owned))
            .transpose()
            .map_err(|_| invalid("a data file name"))?;
        found = Some(name.filter(|name| !name.is_empty()));
    }
    let name = found.ok_or_else(|| invalid("a data entry"))?;
    match name {
        Some(name)
            if !name.contains(['\0', '\\'])
                && std::path::Path::new(&name)
                    .components()
                    .all(|part| matches!(part, std::path::Component::Normal(_))) =>
        {
            Ok(Some(format!("Data/{name}")))
        },
        Some(_) => Err(invalid("a data reference")),
        None => Ok(None),
    }
}

/// The metadata payload without the `DataInfo` of `data`.
pub fn remove_data_info(payload: &[u8], data: u64) -> Result<Vec<u8>> {
    remove_repeated_length_delimited_field_where(payload, METADATA_DATAS_FIELD, |info| {
        Ok::<bool, litchi_iwa_common::Error>(
            view_varint(info, DATA_INFO_IDENTIFIER_FIELD)? == Some(data),
        )
    })
}

/// The data identifiers listed by a `TSP.DataMetadataMap` payload.
pub fn data_metadata_entries(payload: &[u8]) -> Result<Vec<u64>> {
    let mut data = Vec::new();
    for entry in repeated_length_delimited_payloads(payload, 1)? {
        data.push(view_varint(entry, 1)?.ok_or_else(|| invalid("a data entry"))?);
    }
    Ok(data)
}

/// Every `(data, object)` component data-reference owner (current and
/// versioned components) of a metadata payload.
pub fn data_owners(payload: &[u8]) -> Result<Vec<(u64, u64)>> {
    let mut owners = Vec::new();
    for field in [
        METADATA_COMPONENTS_FIELD,
        METADATA_VERSIONED_COMPONENTS_FIELD,
    ] {
        for info in repeated_length_delimited_payloads(payload, field)? {
            for reference in
                repeated_length_delimited_payloads(info, COMPONENT_DATA_REFERENCES_FIELD)?
            {
                let data = view_varint(reference, DATA_REFERENCE_IDENTIFIER_FIELD)?
                    .ok_or_else(|| invalid("a data entry"))?;
                for owner in
                    repeated_length_delimited_payloads(reference, DATA_REFERENCE_OWNERS_FIELD)?
                {
                    let object = view_varint(owner, OWNER_OBJECT_FIELD)?
                        .ok_or_else(|| invalid("a data entry"))?;
                    owners.push((data, object));
                }
            }
        }
    }
    Ok(owners)
}

/// The largest `DataInfo` identifier of a metadata payload (0 for none).
pub fn max_data_identifier(payload: &[u8]) -> Result<u64> {
    let mut maximum = 0;
    for info in repeated_length_delimited_payloads(payload, METADATA_DATAS_FIELD)? {
        maximum = maximum.max(
            view_varint(info, DATA_INFO_IDENTIFIER_FIELD)?
                .ok_or_else(|| invalid("a data entry"))?,
        );
    }
    Ok(maximum)
}

/// The metadata payload with a new `DataInfo` appended (the host's
/// `append_data_info`).
pub fn append_data_info(
    payload: &[u8],
    data: u64,
    digest: &[u8],
    preferred_file_name: &str,
    file_name: &str,
    length: u64,
) -> Result<Vec<u8>> {
    let limits = litchi_iwa_common::WireLimits::default();
    let mut info = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut info, DATA_INFO_IDENTIFIER_FIELD, data)?;
    litchi_iwa_common::wire::append_length_delimited_field(&mut info, 2, digest, limits)?;
    litchi_iwa_common::wire::append_length_delimited_field(
        &mut info,
        3,
        preferred_file_name.as_bytes(),
        limits,
    )?;
    litchi_iwa_common::wire::append_length_delimited_field(
        &mut info,
        DATA_INFO_FILE_NAME_FIELD,
        file_name.as_bytes(),
        limits,
    )?;
    litchi_iwa_common::wire::append_varint_field(&mut info, 18, length)?;
    let mut output = payload.to_vec();
    litchi_iwa_common::wire::append_length_delimited_field(
        &mut output,
        METADATA_DATAS_FIELD,
        &info,
        limits,
    )?;
    Ok(output)
}

/// One new embedded data entry: its preferred file name, SHA-1 digest and
/// length.
#[derive(Debug, Clone, Copy)]
pub struct NewData<'a> {
    /// The caller's preferred file name (`stem.extension`).
    pub preferred_file_name: &'a str,
    /// The SHA-1 digest of the bytes.
    pub digest: &'a [u8],
    /// The byte length.
    pub length: u64,
}

/// The stored file name of data `identifier` (`stem-identifier.extension`,
/// the host's `materialized_file_name`), or `None` for a preferred name
/// that is not one plain file name.
#[must_use]
pub fn materialized_name(preferred: &str, identifier: u64) -> Option<String> {
    let path = std::path::Path::new(preferred);
    if path.file_name().and_then(|name| name.to_str()) != Some(preferred)
        || preferred.contains(['\0', '\\'])
    {
        return None;
    }
    let stem = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.is_empty())?;
    Some(
        match path.extension().and_then(|extension| extension.to_str()) {
            Some(extension) if !extension.is_empty() => format!("{stem}-{identifier}.{extension}"),
            _ => format!("{stem}-{identifier}"),
        },
    )
}

fn wire(error: litchi_iwa_common::Error) -> crate::graph::Error {
    crate::graph::Error::Wire(error)
}

/// Register `new` as the next data entry and make `object` (stored by
/// `member`) its owner, as the host's `insert_unreferenced` followed by
/// `add_component_data_reference` do. Returns the data identifier and the
/// `Data/` member name the caller publishes the bytes under.
///
/// # Errors
///
/// A package without a registry, an object outside every component, an
/// invalid preferred name or registry payload.
pub fn own_new(
    workspace: &mut Workspace<'_>,
    new: NewData<'_>,
    member: &str,
    object: u64,
) -> crate::graph::Result<(u64, String)> {
    let component = workspace
        .component_for_member(member)?
        .ok_or(crate::graph::Error::Invalid(
            "an owner outside every component",
        ))?;
    let payload = workspace
        .registry_payload()?
        .ok_or(crate::graph::Error::Invalid("the package has no registry"))?;
    let data = max_data_identifier(payload)
        .map_err(wire)?
        .checked_add(1)
        .ok_or(crate::graph::Error::IdentifierOverflow)?;
    let stored = materialized_name(new.preferred_file_name, data).ok_or(
        crate::graph::Error::Invalid("an invalid preferred file name"),
    )?;
    let name = format!("Data/{stored}");
    workspace.update_registry(|payload| {
        let payload = append_data_info(
            payload,
            data,
            new.digest,
            new.preferred_file_name,
            &stored,
            new.length,
        )
        .map_err(wire)?;
        adjust_data_reference(&payload, component, data, object, true)
            .map(Some)
            .map_err(wire)
    })?;
    Ok((data, name))
}

/// Drop `object`'s (stored by `member`) owner count of `data`, as the
/// host's `remove_component_data_reference`; data no component references
/// any more loses its `DataInfo` (unless the package keeps a data metadata
/// map), as the host's `remove_orphaned_image_asset`. Returns the `Data/`
/// member the caller deletes.
///
/// # Errors
///
/// A package without a registry, an object outside every component, or no
/// such owner.
pub fn release(
    workspace: &mut Workspace<'_>,
    data: u64,
    member: &str,
    object: u64,
) -> crate::graph::Result<Option<String>> {
    let component = workspace
        .component_for_member(member)?
        .ok_or(crate::graph::Error::Invalid(
            "an owner outside every component",
        ))?;
    // The host's `EmbeddedMediaAsset::is_referenced`: a data metadata map
    // entry or a message-info data reference keeps the asset.
    let mut keep_info = false;
    if let Some(map) = workspace.data_metadata_object()? {
        let object = workspace.object(map)?;
        for message in &object.messages {
            keep_info |= data_metadata_entries(&message.data)
                .map_err(wire)?
                .contains(&data);
        }
    }
    for location in workspace.locations_iter() {
        for object in &workspace.archive_at(location)?.objects {
            keep_info |= object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.data_references.contains(&data));
        }
    }
    let mut deleted = None;
    workspace.update_registry(|payload| {
        let mut payload =
            adjust_data_reference(payload, component, data, object, false).map_err(wire)?;
        if !keep_info && !referenced_data(&payload).map_err(wire)?.contains(&data) {
            deleted = data_member(&payload, data).map_err(wire)?;
            payload = remove_data_info(&payload, data).map_err(wire)?;
        }
        Ok(Some(payload))
    })?;
    Ok(deleted)
}
