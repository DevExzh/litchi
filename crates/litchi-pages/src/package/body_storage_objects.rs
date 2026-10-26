//! Private engine for the objects a rooted text storage owns: the storage
//! itself, the package-metadata watermark and registry edits that creating
//! or deleting such an object requires, and text splices. The ranged-object
//! tables are the shared `litchi_iwa_text_wire::ranged` engine.
//!
//! The rules follow the native graph the migration host writes (record 0371):
//! a new object takes the identifier one above every physical object and
//! every metadata registration, and raises the metadata watermark to it; a
//! deleted object loses its external references and UUID registration, and a
//! deleted watermark identifier releases the suffix to the largest remaining
//! object.

use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_common::wire::{
    WireView, patch_varint_field, remove_repeated_length_delimited_field_where,
    repeated_length_delimited_payloads, transform_length_delimited_fields_at_path,
};
use litchi_iwa_core::{Archive, ArchiveObject};
use litchi_iwa_protos::package_metadata_codec::{
    self, Batch, ComponentDescriptor, ComponentSelector, ExternalReferenceDescriptor,
    ObjectUuidAddition, ObjectUuidDescriptor, PackageMetadataVisitor,
    RewriteError as MetadataError, RewriteOptions, UuidBits,
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
    Ok((storage_at(package, &index, body.root.body)?, budget))
}

/// Resolve one text storage by its native identifier (found through a
/// semantic selector by the caller) and decode its text.
pub(crate) fn storage_by_identifier(package: &Package, identifier: u64) -> GraphResult<Storage> {
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    storage_at(
        package,
        &index,
        NonZeroU64::new(identifier).ok_or(GraphError::InvalidSource)?,
    )
}

fn storage_at(
    package: &Package,
    index: &ObjectIndex<'_>,
    identifier: NonZeroU64,
) -> GraphResult<Storage> {
    let located = index.get(identifier)?;
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
    Ok(Storage {
        identifier,
        component: index.component_name(located.component_index)?,
        message_index,
        message_type: message.type_,
        payload: message.data.clone(),
        text: text.text().to_owned(),
    })
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
    uuid_bits: std::collections::HashSet<(u64, u64)>,
    /// The object holding the `TSP.DataMetadataMap`, if any.
    pub(crate) data_metadata_map: Option<u64>,
}

#[derive(Default)]
struct FactsVisitor {
    max: u64,
    components: Vec<(u64, String, bool)>,
    uuids: Vec<(u64, u64)>,
    uuid_bits: std::collections::HashSet<(u64, u64)>,
    data_metadata_map: Option<u64>,
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
        self.uuid_bits
            .insert((binding.uuid().lower(), binding.uuid().upper()));
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
        self.data_metadata_map = Some(object_identifier);
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
        uuid_bits: visitor.uuid_bits,
        data_metadata_map: visitor.data_metadata_map,
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

impl Metadata {
    /// The payload with fresh UUID registrations for `objects` in the
    /// current component `component` and the watermark raised to
    /// `new_last` (the host's `add_component_object_uuids` followed by
    /// `set_package_last_object_identifier`). The leaf draws deterministic
    /// UUIDs, unique within the package, where the host draws random ones.
    pub(crate) fn registered(
        &self,
        package: &Package,
        component: u64,
        objects: &[u64],
        new_last: u64,
    ) -> GraphResult<Vec<u8>> {
        let locator = self
            .components
            .iter()
            .filter(|(identifier, _, current)| *current && *identifier == component)
            .map(|(_, locator, _)| locator.as_str())
            .collect::<Vec<_>>();
        let [locator] = locator.as_slice() else {
            return Err(GraphError::InvalidSource);
        };
        if objects.iter().any(|object| {
            self.uuids
                .iter()
                .any(|(_, registered)| registered == object)
        }) {
            return Err(GraphError::InvalidSource);
        }
        let selector = ComponentSelector::new(component, locator);
        let mut taken = self.uuid_bits.clone();
        let mut additions = Vec::with_capacity(objects.len());
        for object in objects {
            let uuid = fresh_uuid(&taken, *object);
            taken.insert((uuid.lower(), uuid.upper()));
            additions.push(ObjectUuidAddition::new(selector, *object, uuid));
        }
        package_metadata_codec::rewrite_package_metadata(
            &self.payload,
            Batch::new(self.last, new_last.max(self.last), &additions, &[]),
            metadata_options(package, self.payload.len())?,
        )
        .map(package_metadata_codec::RewriteOutput::into_bytes)
        .map_err(|_| GraphError::InvalidSource)
    }
}

/// A deterministic UUID for a created object, distinct from every `taken`
/// pair and never nil.
fn fresh_uuid(taken: &std::collections::HashSet<(u64, u64)>, identifier: u64) -> UuidBits {
    let mut lower = identifier.max(1);
    let mut upper = identifier.rotate_left(17) ^ 0x9e37_79b9_7f4a_7c15;
    while taken.contains(&(lower, upper)) {
        lower = lower.wrapping_add(1).max(1);
        upper = upper.rotate_left(7) ^ 0xd1b5_4a32_d192_ed03;
    }
    UuidBits::new(lower, upper)
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

/// Replace the UTF-16 `start..end` of a storage's text through the shared
/// TSWP rewrite (the host's `replace_storage_text` splice), returning the
/// object references the splice removed.
pub(crate) fn splice_text(
    package: &Package,
    storage: &mut Storage,
    start: u32,
    end: u32,
    replacement: &str,
) -> GraphResult<Vec<u64>> {
    if start > end || end > storage.utf16_len() || !scalar_boundaries(&storage.text, [start, end]) {
        return Err(GraphError::InvalidSource);
    }
    let limits = super::storage_rewrite_limits(package.state.source.limits())
        .map_err(|_| GraphError::InvalidSource)?;
    let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
        &storage.payload,
        start as usize..end as usize,
        replacement,
        litchi_iwa_text_wire::RewriteBehavior::ReplaceSelection,
        limits,
    )
    .map_err(|_| GraphError::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    let rewritten = prepared
        .execute(litchi_iwa_text_wire::StorageRewriteExecutionLimits {
            max_output_bytes: requirements.output_bytes(),
            max_retained_elements: requirements.retained_elements(),
            max_retained_bytes: requirements.retained_bytes(),
            max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
            max_allocations: requirements.allocations(),
            max_work: requirements.work(),
        })
        .map_err(|_| GraphError::InvalidSource)?;
    let removed = rewritten.removed_object_references().to_vec();
    let byte = |units: u32| {
        let mut count = 0u32;
        for (index, character) in storage.text.char_indices() {
            if count == units {
                return index;
            }
            count += character.len_utf16() as u32;
        }
        storage.text.len()
    };
    let (from, to) = (byte(start), byte(end));
    let mut text = String::with_capacity(storage.text.len() + replacement.len());
    text.push_str(&storage.text[..from]);
    text.push_str(replacement);
    text.push_str(&storage.text[to..]);
    storage.payload = rewritten.into_bytes();
    storage.text = text;
    Ok(removed)
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

fn view_varint(payload: &[u8], number: u32) -> litchi_iwa_common::Result<Option<u64>> {
    let view = WireView::parse(payload, litchi_iwa_common::WireLimits::default())?;
    Ok(varint_of(&view, number))
}

/// Add one count to, or remove one count from, `component`'s
/// data-reference owner (`data`, `object`) in a metadata payload (the host's
/// `{add,remove}_component_data_reference`): a removal that empties the
/// only owner drops the whole data reference, one that empties an owner
/// drops the owner, and an addition appends a new reference or owner when
/// none exists.
pub(crate) fn adjust_data_reference(
    payload: &[u8],
    component: u64,
    data: u64,
    object: u64,
    add: bool,
) -> GraphResult<Vec<u8>> {
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
        })
        .map_err(map_wire_error)?;
    }
    if matched != 1 {
        return Err(GraphError::InvalidSource);
    }
    Ok(output)
}

fn adjust_component(
    info: &[u8],
    data: u64,
    object: u64,
    add: bool,
) -> litchi_iwa_common::Result<Vec<u8>> {
    let invalid = || litchi_iwa_common::Error::InvalidFormat("data reference".to_owned());
    let references = repeated_length_delimited_payloads(info, COMPONENT_DATA_REFERENCES_FIELD)?;
    let mut matches = Vec::new();
    for reference in references {
        if view_varint(reference, DATA_REFERENCE_IDENTIFIER_FIELD)? == Some(data) {
            matches.push(reference);
        }
    }
    let is_data = |reference: &[u8]| -> litchi_iwa_common::Result<bool> {
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

fn adjust_owner(
    reference: &[u8],
    object: u64,
    add: bool,
    found: &[u64],
) -> litchi_iwa_common::Result<Vec<u8>> {
    let is_owner = |owner: &[u8]| -> litchi_iwa_common::Result<bool> {
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
pub(crate) fn referenced_data(payload: &[u8]) -> GraphResult<std::collections::HashSet<u64>> {
    let mut referenced = std::collections::HashSet::new();
    for field in [
        METADATA_COMPONENTS_FIELD,
        METADATA_VERSIONED_COMPONENTS_FIELD,
    ] {
        for info in repeated_length_delimited_payloads(payload, field).map_err(map_wire_error)? {
            for reference in
                repeated_length_delimited_payloads(info, COMPONENT_DATA_REFERENCES_FIELD)
                    .map_err(map_wire_error)?
            {
                if let Some(data) = view_varint(reference, DATA_REFERENCE_IDENTIFIER_FIELD)
                    .map_err(map_wire_error)?
                {
                    referenced.insert(data);
                }
            }
        }
    }
    Ok(referenced)
}

/// The `Data/` member name of a data entry, from its `DataInfo.file_name`.
pub(crate) fn data_member(payload: &[u8], data: u64) -> GraphResult<Option<String>> {
    let mut found = None;
    for info in
        repeated_length_delimited_payloads(payload, METADATA_DATAS_FIELD).map_err(map_wire_error)?
    {
        if view_varint(info, DATA_INFO_IDENTIFIER_FIELD).map_err(map_wire_error)? != Some(data) {
            continue;
        }
        if found.is_some() {
            return Err(GraphError::InvalidSource);
        }
        let view = WireView::parse(info, litchi_iwa_common::WireLimits::default())
            .map_err(map_wire_error)?;
        let name = view
            .fields()
            .find(|field| field.number() == DATA_INFO_FILE_NAME_FIELD)
            .map(|field| std::str::from_utf8(field.payload()).map(str::to_owned))
            .transpose()
            .map_err(|_| GraphError::InvalidSource)?;
        found = Some(name.filter(|name| !name.is_empty()));
    }
    let name = found.ok_or(GraphError::InvalidSource)?;
    match name {
        Some(name)
            if !name.contains(['\0', '\\'])
                && std::path::Path::new(&name)
                    .components()
                    .all(|part| matches!(part, std::path::Component::Normal(_))) =>
        {
            Ok(Some(format!("Data/{name}")))
        },
        Some(_) => Err(GraphError::InvalidSource),
        None => Ok(None),
    }
}

/// The metadata payload without the `DataInfo` of `data`.
pub(crate) fn remove_data_info(payload: &[u8], data: u64) -> GraphResult<Vec<u8>> {
    remove_repeated_length_delimited_field_where(payload, METADATA_DATAS_FIELD, |info| {
        Ok::<bool, litchi_iwa_common::Error>(
            view_varint(info, DATA_INFO_IDENTIFIER_FIELD)? == Some(data),
        )
    })
    .map_err(map_wire_error)
}

/// The data identifiers listed by a `TSP.DataMetadataMap` payload.
pub(crate) fn data_metadata_entries(payload: &[u8]) -> GraphResult<Vec<u64>> {
    let mut data = Vec::new();
    for entry in repeated_length_delimited_payloads(payload, 1).map_err(map_wire_error)? {
        data.push(
            view_varint(entry, 1)
                .map_err(map_wire_error)?
                .ok_or(GraphError::InvalidSource)?,
        );
    }
    Ok(data)
}
