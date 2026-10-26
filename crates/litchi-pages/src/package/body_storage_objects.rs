//! Private engine for the objects a rooted text storage owns: its ranged
//! object tables (bookmarks, smart fields), the package-metadata watermark
//! and registry edits that creating or deleting such an object requires.
//!
//! The rules follow the native graph the migration host writes (record 0371):
//! a new object takes the identifier one above every physical object and
//! every metadata registration, and raises the metadata watermark to it; a
//! deleted object loses its external references and UUID registration, and a
//! deleted watermark identifier releases the suffix to the largest remaining
//! object. Table entries keep their raw bytes; new entries are canonical.

use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_common::wire::{
    WireView, patch_length_delimited_field, patch_varint_field,
    remove_repeated_length_delimited_field_where, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_fields_at_path,
};
use litchi_iwa_core::{Archive, ArchiveObject};
use litchi_iwa_protos::package_metadata_codec::{
    self, Batch, ComponentDescriptor, ExternalReferenceDescriptor, ObjectUuidDescriptor,
    PackageMetadataVisitor, RewriteError as MetadataError, RewriteOptions,
};

use super::Package;
use super::body_drawable_graph::{
    self as graph, Budget, GraphError, GraphResult, ObjectIndex, map_archive_error, map_wire_error,
};

const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const METADATA_LAST_IDENTIFIER_FIELD: u32 = 1;
const METADATA_COMPONENTS_FIELD: u32 = 3;
const COMPONENT_IDENTIFIER_FIELD: u32 = 1;
const COMPONENT_EXTERNAL_REFERENCES_FIELD: u32 = 6;
const COMPONENT_OBJECT_UUIDS_FIELD: u32 = 11;
const EXTERNAL_COMPONENT_FIELD: u32 = 1;
const EXTERNAL_OBJECT_FIELD: u32 = 2;
const UUID_ENTRY_IDENTIFIER_FIELD: u32 = 1;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;

/// The rooted body text storage.
pub(crate) struct Storage {
    pub(crate) identifier: NonZeroU64,
    pub(crate) component: Arc<str>,
    pub(crate) message_index: usize,
    pub(crate) message_type: u32,
    pub(crate) payload: Vec<u8>,
    pub(crate) text: String,
}

impl Storage {
    /// The text length in UTF-16 code units.
    pub(crate) fn utf16_len(&self) -> u32 {
        self.text.encode_utf16().count() as u32
    }
}

/// Resolve the rooted body storage and decode its text.
pub(crate) fn body_storage(package: &Package) -> GraphResult<(Storage, Budget)> {
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let located = index.get(body.root.body)?;
    let mut selected = None;
    for (message_index, message) in located.object.messages.iter().enumerate() {
        if matches!(message.type_, 2_001 | 2_022)
            && selected.replace((message_index, message)).is_some()
        {
            return Err(GraphError::InvalidSource);
        }
    }
    let (message_index, message) = selected.ok_or(GraphError::InvalidSource)?;
    graph::validate_message_metadata(located.object, message_index)?;
    let limits = super::storage_rewrite_limits(package.state.source.limits())
        .map_err(|_| GraphError::InvalidSource)?;
    let text_limits = litchi_iwa_text_wire::Limits::new(
        limits.max_message_bytes().max(message.data.len()).max(1),
        limits.max_fields(),
        limits.max_fragments(),
        limits.max_text_bytes(),
    )
    .map_err(|_| GraphError::InvalidSource)?;
    let text = litchi_iwa_text_wire::from_bytes(&message.data, text_limits)
        .map_err(|_| GraphError::InvalidSource)?;
    Ok((
        Storage {
            identifier: body.root.body,
            component: index.component_name(located.component_index)?,
            message_index,
            message_type: message.type_,
            payload: message.data.clone(),
            text: text.text().to_owned(),
        },
        budget,
    ))
}

/// One boundary of an `ObjectAttributeTable`: the object that starts at a
/// UTF-16 index (or `None`, a gap), with its raw entry bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Boundary {
    pub(crate) index: u32,
    pub(crate) object: Option<u64>,
    raw: Vec<u8>,
}

/// The raw ranged table `field` of a storage payload and its boundaries.
pub(crate) fn ranged_table(
    storage: &Storage,
    field: u32,
) -> GraphResult<(Option<Vec<u8>>, Vec<Boundary>)> {
    let tables =
        repeated_length_delimited_payloads(&storage.payload, field).map_err(map_wire_error)?;
    let table = match tables.as_slice() {
        [] => return Ok((None, Vec::new())),
        [table] => table.to_vec(),
        _ => return Err(GraphError::InvalidSource),
    };
    let entries =
        repeated_length_delimited_payloads(&table, TABLE_ENTRIES_FIELD).map_err(map_wire_error)?;
    if entries.is_empty() {
        return Err(GraphError::InvalidSource);
    }
    let mut boundaries = Vec::with_capacity(entries.len());
    for raw in entries {
        let view = WireView::parse(raw, litchi_iwa_common::WireLimits::default())
            .map_err(map_wire_error)?;
        let index =
            graph::unique_field(&view, ENTRY_INDEX_FIELD, 0)?.ok_or(GraphError::InvalidSource)?;
        let index = u32::try_from(graph::canonical_varint(index.payload())?)
            .map_err(|_| GraphError::InvalidSource)?;
        let object = graph::unique_field(&view, ENTRY_OBJECT_FIELD, 2)?
            .map(|field| reference_identifier(field.payload()))
            .transpose()?;
        boundaries.push(Boundary {
            index,
            object,
            raw: raw.to_vec(),
        });
    }
    validate_boundaries(&boundaries, storage)?;
    let table = Some(table);
    Ok((table, boundaries))
}

fn reference_identifier(payload: &[u8]) -> GraphResult<u64> {
    let view = WireView::parse(payload, litchi_iwa_common::WireLimits::default())
        .map_err(map_wire_error)?;
    let field = graph::unique_field(&view, 1, 0)?.ok_or(GraphError::InvalidSource)?;
    let identifier = graph::canonical_varint(field.payload())?;
    if identifier == 0 {
        return Err(GraphError::InvalidSource);
    }
    Ok(identifier)
}

fn validate_boundaries(boundaries: &[Boundary], storage: &Storage) -> GraphResult<()> {
    let Some(first) = boundaries.first() else {
        return Ok(());
    };
    if first.index != 0 {
        return Err(GraphError::InvalidSource);
    }
    let length = storage.utf16_len();
    for pair in boundaries.windows(2) {
        if pair[0].index >= pair[1].index {
            return Err(GraphError::InvalidSource);
        }
    }
    if boundaries
        .iter()
        .any(|boundary| boundary.object.is_some() && boundary.index == length)
    {
        return Err(GraphError::InvalidSource);
    }
    if !scalar_boundaries(
        &storage.text,
        boundaries.iter().map(|boundary| boundary.index),
    ) {
        return Err(GraphError::InvalidSource);
    }
    Ok(())
}

/// Whether every sorted UTF-16 index falls on a scalar boundary of `text`.
pub(crate) fn scalar_boundaries(text: &str, indexes: impl IntoIterator<Item = u32>) -> bool {
    let mut indexes = indexes.into_iter().peekable();
    let mut index = 0u32;
    while indexes.peek() == Some(&index) {
        indexes.next();
    }
    for character in text.chars() {
        index += character.len_utf16() as u32;
        while indexes.peek() == Some(&index) {
            indexes.next();
        }
        if indexes.peek().is_some_and(|next| *next < index) {
            return false;
        }
    }
    indexes.next().is_none()
}

/// The ranges `[start, end)` each object of the table covers.
pub(crate) fn ranges(boundaries: &[Boundary], length: u32) -> Vec<(u64, u32, u32)> {
    boundaries
        .iter()
        .enumerate()
        .filter_map(|(position, boundary)| {
            let object = boundary.object?;
            let end = boundaries
                .get(position + 1)
                .map_or(length, |next| next.index);
            Some((object, boundary.index, end))
        })
        .collect()
}

/// Whether `[start, end)` overlaps an object's range other than `ignored`.
pub(crate) fn overlaps(
    boundaries: &[Boundary],
    length: u32,
    start: u32,
    end: u32,
    ignored: Option<u64>,
) -> bool {
    ranges(boundaries, length)
        .into_iter()
        .any(|(object, from, to)| Some(object) != ignored && start < to && from < end)
}

/// Add `[start, end)` for `object` (the host's `add_range`).
pub(crate) fn add_range(
    boundaries: &mut Vec<Boundary>,
    start: u32,
    end: u32,
    object: u64,
) -> GraphResult<()> {
    if boundaries.is_empty() {
        boundaries.push(new_boundary(0, None));
    }
    set_boundary(boundaries, start, Some(object))?;
    set_boundary(boundaries, end, None)?;
    boundaries.dedup_by(|current, previous| current.object == previous.object);
    Ok(())
}

/// Remove the range of `object` (the host's `remove_range`).
pub(crate) fn remove_range(boundaries: &mut Vec<Boundary>, object: u64) -> GraphResult<()> {
    let matches = boundaries
        .iter()
        .enumerate()
        .filter(|(_, boundary)| boundary.object == Some(object))
        .map(|(position, _)| position)
        .collect::<Vec<_>>();
    let [position] = matches.as_slice() else {
        return Err(GraphError::InvalidSource);
    };
    let boundary = &mut boundaries[*position];
    boundary.raw = patch_entry_object(&boundary.raw, boundary.object, None)?;
    boundary.object = None;
    boundaries.dedup_by(|current, previous| current.object == previous.object);
    Ok(())
}

fn set_boundary(
    boundaries: &mut Vec<Boundary>,
    index: u32,
    object: Option<u64>,
) -> GraphResult<()> {
    if let Some(boundary) = boundaries
        .iter_mut()
        .find(|boundary| boundary.index == index)
    {
        if boundary.object != object {
            boundary.raw = patch_entry_object(&boundary.raw, boundary.object, object)?;
            boundary.object = object;
        }
        return Ok(());
    }
    boundaries.push(new_boundary(index, object));
    boundaries.sort_by_key(|boundary| boundary.index);
    Ok(())
}

fn encode_reference(identifier: u64) -> Vec<u8> {
    let mut reference = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut reference, 1, identifier).unwrap_or_default();
    reference
}

fn new_boundary(index: u32, object: Option<u64>) -> Boundary {
    let mut raw = Vec::new();
    let _ =
        litchi_iwa_common::wire::append_varint_field(&mut raw, ENTRY_INDEX_FIELD, u64::from(index));
    if let Some(object) = object {
        let _ = litchi_iwa_common::wire::append_length_delimited_field(
            &mut raw,
            ENTRY_OBJECT_FIELD,
            &encode_reference(object),
            litchi_iwa_common::WireLimits::default(),
        );
    }
    Boundary { index, object, raw }
}

fn patch_entry_object(
    raw: &[u8],
    current: Option<u64>,
    replacement: Option<u64>,
) -> GraphResult<Vec<u8>> {
    let reference = replacement.map(encode_reference);
    patch_length_delimited_field(
        raw,
        ENTRY_OBJECT_FIELD,
        current.is_some(),
        reference.as_deref(),
    )
    .map_err(map_wire_error)
}

/// The table payload for `boundaries`; `None` when no object remains.
pub(crate) fn encode_table(
    original: Option<&[u8]>,
    boundaries: &[Boundary],
) -> GraphResult<Vec<u8>> {
    let entries = boundaries
        .iter()
        .map(|boundary| boundary.raw.clone())
        .collect::<Vec<_>>();
    match original {
        Some(table) => {
            rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &entries)
                .map_err(map_wire_error)
        },
        None => {
            let mut table = Vec::new();
            for entry in &entries {
                litchi_iwa_common::wire::append_length_delimited_field(
                    &mut table,
                    TABLE_ENTRIES_FIELD,
                    entry,
                    litchi_iwa_common::WireLimits::default(),
                )
                .map_err(map_wire_error)?;
            }
            Ok(table)
        },
    }
}

/// Replace (or remove, with `None`) the ranged table `field` of a storage
/// payload.
pub(crate) fn replace_table(
    payload: &[u8],
    field: u32,
    present: bool,
    table: Option<&[u8]>,
) -> GraphResult<Vec<u8>> {
    patch_length_delimited_field(payload, field, present, table).map_err(map_wire_error)
}

/// The package metadata component and its facts.
pub(crate) struct Metadata {
    pub(crate) component: Arc<str>,
    pub(crate) object: u64,
    pub(crate) message_index: usize,
    pub(crate) payload: Vec<u8>,
    pub(crate) last: u64,
    max_registered: u64,
    components: Vec<(u64, String, bool)>,
    uuids: Vec<(u64, u64)>,
}

#[derive(Default)]
struct FactsVisitor {
    max: u64,
    components: Vec<(u64, String, bool)>,
    uuids: Vec<(u64, u64)>,
}

impl PackageMetadataVisitor for FactsVisitor {
    fn visit_component(&mut self, component: ComponentDescriptor<'_>) -> Result<(), MetadataError> {
        self.max = self.max.max(component.identifier());
        self.components.push((
            component.identifier(),
            component.effective_locator().to_owned(),
            component.is_current(),
        ));
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataError> {
        self.max = self.max.max(binding.object_identifier());
        if binding.component().is_current() {
            self.uuids.push((
                binding.component().identifier(),
                binding.object_identifier(),
            ));
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataError> {
        if let Some(identifier) = reference.object_identifier() {
            self.max = self.max.max(identifier);
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), MetadataError> {
        self.max = self.max.max(owner.object_identifier());
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), MetadataError> {
        self.max = self.max.max(identifier);
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), MetadataError> {
        self.max = self.max.max(object_identifier);
        Ok(())
    }
}

fn metadata_options(package: &Package, input: usize) -> GraphResult<RewriteOptions> {
    let archive_limits = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let input = input.max(1);
    let maximum = package
        .state
        .source
        .limits()
        .max_iwa_stream_bytes()
        .max(input);
    Ok(RewriteOptions::new(
        input,
        maximum,
        archive_limits.max_header_fields().max(input.min(1_000_000)),
        input.saturating_mul(128).max(1),
        u32::try_from(archive_limits.max_header_nesting()).unwrap_or(u32::MAX),
        input,
        input,
        64,
    ))
}

/// Read the package metadata, if the package has the sidecar.
pub(crate) fn metadata(package: &Package) -> GraphResult<Option<Metadata>> {
    let Some(component) = package.state.source.components().get(METADATA_COMPONENT) else {
        return Ok(None);
    };
    let mut location = None;
    for object in &component.archive().objects {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == METADATA_MESSAGE_TYPE
                && location
                    .replace((object.archive_info.identifier, message_index, message))
                    .is_some()
            {
                return Err(GraphError::InvalidSource);
            }
        }
    }
    let (object, message_index, message) = location.ok_or(GraphError::InvalidSource)?;
    let mut visitor = FactsVisitor::default();
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        &message.data,
        metadata_options(package, message.data.len())?,
        &mut visitor,
    )
    .map_err(|_| GraphError::InvalidSource)?;
    Ok(Some(Metadata {
        component: Arc::from(component.name()),
        object: object.ok_or(GraphError::InvalidSource)?,
        message_index,
        payload: message.data.clone(),
        last: inspection.last_object_identifier(),
        max_registered: visitor.max,
        components: visitor.components,
        uuids: visitor.uuids,
    }))
}

/// The identifier a new object takes: one above every physical object, the
/// metadata watermark and every metadata registration.
pub(crate) fn next_identifier(package: &Package, metadata: Option<&Metadata>) -> GraphResult<u64> {
    let physical = package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .filter_map(|object| object.archive_info.identifier)
        .max()
        .unwrap_or(0);
    let registered = metadata.map_or(0, |metadata| metadata.last.max(metadata.max_registered));
    physical
        .max(registered)
        .checked_add(1)
        .ok_or(GraphError::InvalidSource)
}

impl Metadata {
    /// The current component whose locator is the archive `entry`
    /// (`Index/<locator>.iwa`).
    pub(crate) fn component_for_entry(&self, entry: &str) -> GraphResult<Option<u64>> {
        let locator = entry
            .strip_prefix("Index/")
            .and_then(|name| name.strip_suffix(".iwa"))
            .ok_or(GraphError::InvalidSource)?;
        let matches = self
            .components
            .iter()
            .filter(|(_, effective, current)| *current && effective == locator)
            .map(|(identifier, _, _)| *identifier)
            .collect::<Vec<_>>();
        let mut matches = matches;
        matches.sort_unstable();
        matches.dedup();
        match matches.as_slice() {
            [] => Ok(None),
            [identifier] => Ok(Some(*identifier)),
            _ => Err(GraphError::InvalidSource),
        }
    }

    /// The current component that registers a UUID for `object`.
    pub(crate) fn uuid_owner(&self, object: u64) -> GraphResult<Option<u64>> {
        let mut owners = self
            .uuids
            .iter()
            .filter(|(_, registered)| *registered == object)
            .map(|(component, _)| *component)
            .collect::<Vec<_>>();
        owners.dedup();
        match owners.as_slice() {
            [] => Ok(None),
            [component] => Ok(Some(*component)),
            _ => Err(GraphError::InvalidSource),
        }
    }

    /// The payload with the watermark raised to `identifier`.
    pub(crate) fn raised(&self, package: &Package, identifier: u64) -> GraphResult<Vec<u8>> {
        if identifier == self.last {
            return Ok(self.payload.clone());
        }
        package_metadata_codec::rewrite_package_metadata(
            &self.payload,
            Batch::new(self.last, identifier, &[], &[]),
            metadata_options(package, self.payload.len())?,
        )
        .map(package_metadata_codec::RewriteOutput::into_bytes)
        .map_err(|_| GraphError::InvalidSource)
    }
}

/// Remove `component`'s external references to `object` from a metadata
/// payload (the host's `remove_component_external_references_to_object`).
pub(crate) fn remove_external_references(
    payload: &[u8],
    component: u64,
    object: u64,
) -> GraphResult<Vec<u8>> {
    let matches = |reference: &[u8]| -> litchi_iwa_common::Result<bool> {
        let view = WireView::parse(reference, litchi_iwa_common::WireLimits::default())?;
        Ok(
            varint_of(&view, EXTERNAL_COMPONENT_FIELD) == Some(component)
                && varint_of(&view, EXTERNAL_OBJECT_FIELD) == Some(object),
        )
    };
    transform_length_delimited_fields_at_path(payload, &[METADATA_COMPONENTS_FIELD], |data| {
        let references =
            repeated_length_delimited_payloads(data, COMPONENT_EXTERNAL_REFERENCES_FIELD)?;
        let mut count = 0usize;
        for reference in references {
            count += usize::from(matches(reference)?);
        }
        match count {
            0 => Ok(data.to_vec()),
            1 => remove_repeated_length_delimited_field_where(
                data,
                COMPONENT_EXTERNAL_REFERENCES_FIELD,
                matches,
            ),
            _ => Err(litchi_iwa_common::Error::InvalidFormat(
                "duplicate external reference".to_owned(),
            )),
        }
    })
    .map_err(map_wire_error)
}

fn varint_of(view: &WireView<'_>, number: u32) -> Option<u64> {
    view.fields()
        .find(|field| field.number() == number)
        .and_then(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
        .map(|(value, _)| value)
}

/// Remove `object`'s UUID registration from `component` (the host's
/// `remove_component_object_uuids`).
pub(crate) fn remove_object_uuid(
    payload: &[u8],
    component: u64,
    object: u64,
) -> GraphResult<Vec<u8>> {
    transform_length_delimited_fields_at_path(payload, &[METADATA_COMPONENTS_FIELD], |data| {
        let view = WireView::parse(data, litchi_iwa_common::WireLimits::default())?;
        let identifier = view
            .fields()
            .find(|field| field.number() == COMPONENT_IDENTIFIER_FIELD)
            .and_then(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
            .map(|(value, _)| value);
        if identifier != Some(component) {
            return Ok(data.to_vec());
        }
        remove_repeated_length_delimited_field_where(data, COMPONENT_OBJECT_UUIDS_FIELD, |entry| {
            let view = WireView::parse(entry, litchi_iwa_common::WireLimits::default())?;
            Ok::<bool, litchi_iwa_common::Error>(
                view.fields()
                    .find(|field| field.number() == UUID_ENTRY_IDENTIFIER_FIELD)
                    .and_then(|field| {
                        litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok()
                    })
                    .is_some_and(|(value, _)| value == object),
            )
        })
    })
    .map_err(map_wire_error)
}

/// The metadata payload after deleting physical objects whose largest
/// identifier may have been the watermark (the host's
/// `release_package_identifier_suffix`): `remaining_max` is the largest
/// identifier left in the candidate.
pub(crate) fn released(
    payload: &[u8],
    last: u64,
    removed: &[u64],
    remaining_max: u64,
) -> GraphResult<Vec<u8>> {
    if !removed.contains(&last) {
        return Ok(payload.to_vec());
    }
    if remaining_max >= last {
        return Err(GraphError::InvalidSource);
    }
    patch_varint_field(
        payload,
        METADATA_LAST_IDENTIFIER_FIELD,
        true,
        Some(remaining_max),
    )
    .map_err(map_wire_error)
}

/// The largest object identifier in `archives` (staged) and every other
/// component of `package`.
pub(crate) fn max_identifier(package: &Package, archives: &[(Arc<str>, Archive)]) -> u64 {
    package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| {
            let staged = archives
                .iter()
                .find(|(name, _)| name.as_ref() == component.name())
                .map_or(component.archive(), |(_, archive)| archive);
            staged.objects.iter()
        })
        .filter_map(|object| object.archive_info.identifier)
        .max()
        .unwrap_or(0)
}

/// A new single-message object.
pub(crate) fn new_object(
    package: &Package,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
) -> GraphResult<ArchiveObject> {
    let limits = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    ArchiveObject::new(
        identifier,
        vec![litchi_iwa_core::RawMessage {
            type_: message_type,
            data,
        }],
        limits,
    )
    .map_err(graph::map_core_error)
}

/// A random lower-case hyphenated UUID for a smart field's
/// `text_attribute_uuid_string`.
pub(crate) fn text_attribute_uuid() -> GraphResult<String> {
    let braced = litchi_core::id::generate_guid_braced();
    braced
        .strip_prefix('{')
        .and_then(|uuid| uuid.strip_suffix('}'))
        .map(str::to_owned)
        .ok_or(GraphError::InvalidSource)
}

/// Whether a smart field's UUID string is the canonical 8-4-4-4-12 hex form.
pub(crate) fn valid_text_attribute_uuid(uuid: &str) -> bool {
    uuid.len() == 36
        && uuid.bytes().enumerate().all(|(index, byte)| match index {
            8 | 13 | 18 | 23 => byte == b'-',
            _ => byte.is_ascii_hexdigit(),
        })
}
