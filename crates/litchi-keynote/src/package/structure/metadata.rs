//! Raw-preserving edits of the package registry (`TSP.PackageMetadata`).
//!
//! The registry lists every component (`ComponentInfo`), its external
//! references, data references and object UUIDs, the package object
//! watermark and the save tokens. Structural edits append components and
//! registry records, remove them, and advance the watermark and the save
//! tokens of every component they touch. Untouched components keep their
//! exact bytes; touched ones keep every field this editor does not model.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::varint::{decode_varint_from_bytes, encode_varint_into};
use litchi_iwa_common::wire::{WireField, parse_wire_fields};

use super::StructureError;

const LAST_OBJECT_IDENTIFIER: u32 = 1;
const COMPONENTS: u32 = 3;
const DATAS: u32 = 4;
const ROOT_SAVE_TOKEN: u32 = 8;

const COMPONENT_IDENTIFIER: u32 = 1;
const COMPONENT_PREFERRED_LOCATOR: u32 = 2;
const COMPONENT_LOCATOR: u32 = 3;
const COMPONENT_EXTERNAL_REFERENCES: u32 = 6;
const COMPONENT_DATA_REFERENCES: u32 = 7;
const COMPONENT_OBJECT_UUIDS: u32 = 11;
const COMPONENT_SAVE_TOKEN: u32 = 12;
const COMPONENT_VERSIONED_EXTERNAL_REFERENCES: u32 = 18;
const COMPONENT_AMBIGUOUS_IDENTIFIERS: u32 = 20;

const VARINT: u8 = 0;
const LENGTH_DELIMITED: u8 = 2;

/// A 128-bit native UUID (`TSP.UUID { lower = 1; upper = 2; }`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(in crate::package) struct Uuid {
    pub(in crate::package) lower: u64,
    pub(in crate::package) upper: u64,
}

/// One `ComponentExternalReference`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::package) struct ExternalReference {
    pub(in crate::package) component: u64,
    pub(in crate::package) object: Option<u64>,
    pub(in crate::package) is_weak: Option<bool>,
}

/// One `ComponentDataReference`: a data item and the objects that own it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::package) struct DataReference {
    pub(in crate::package) data: u64,
    pub(in crate::package) owners: Vec<(u64, u32)>,
}

/// A decoded, editable `ComponentInfo`. Fields this editor does not model
/// are kept as raw encoded fields, in source order.
#[derive(Debug, Clone)]
pub(in crate::package) struct Component {
    pub(in crate::package) identifier: u64,
    pub(in crate::package) locator: String,
    external_references: Vec<ExternalReference>,
    data_references: Vec<DataReference>,
    object_uuids: Vec<(u64, Uuid)>,
    save_token: Option<u64>,
    /// Every other field, raw, in source order (identifier and locators
    /// included, so they keep their exact encoding).
    other: Vec<Vec<u8>>,
}

impl Component {
    /// The component's effective package member name.
    pub(in crate::package) fn member_name(&self) -> String {
        format!("Index/{}.iwa", self.locator)
    }

    pub(in crate::package) fn external_references(&self) -> &[ExternalReference] {
        &self.external_references
    }

    pub(in crate::package) fn object_uuids(&self) -> &[(u64, Uuid)] {
        &self.object_uuids
    }

    pub(in crate::package) fn uuid_of(&self, object: u64) -> Option<Uuid> {
        self.object_uuids
            .iter()
            .find(|(identifier, _)| *identifier == object)
            .map(|(_, uuid)| *uuid)
    }

    pub(in crate::package) fn add_object_uuid(&mut self, object: u64, uuid: Uuid) {
        self.object_uuids.push((object, uuid));
    }

    pub(in crate::package) fn add_external_reference(&mut self, reference: ExternalReference) {
        if !self.external_references.contains(&reference) {
            self.external_references.push(reference);
        }
    }

    pub(in crate::package) fn add_data_owner(&mut self, data: u64, object: u64) {
        if let Some(entry) = self
            .data_references
            .iter_mut()
            .find(|entry| entry.data == data)
        {
            if let Some(owner) = entry.owners.iter_mut().find(|(owner, _)| *owner == object) {
                owner.1 = owner.1.saturating_add(1);
            } else {
                entry.owners.push((object, 1));
            }
        } else {
            self.data_references.push(DataReference {
                data,
                owners: vec![(object, 1)],
            });
        }
    }

    /// Remove the UUID records of `objects`; returns how many were removed.
    pub(in crate::package) fn remove_object_uuids(&mut self, objects: &[u64]) -> usize {
        let before = self.object_uuids.len();
        self.object_uuids
            .retain(|(identifier, _)| !objects.contains(identifier));
        before - self.object_uuids.len()
    }

    /// The objects that own data item `data`.
    pub(in crate::package) fn data_owned_objects(&self, data: u64) -> Vec<u64> {
        self.data_references
            .iter()
            .filter(|entry| entry.data == data)
            .flat_map(|entry| entry.owners.iter().map(|(object, _)| *object))
            .collect()
    }

    /// The data items any of `objects` owns, ascending and unique.
    pub(in crate::package) fn data_owned_by(&self, objects: &[u64]) -> Vec<u64> {
        let mut data = self
            .data_references
            .iter()
            .filter(|entry| {
                entry
                    .owners
                    .iter()
                    .any(|(owner, _)| objects.contains(owner))
            })
            .map(|entry| entry.data)
            .collect::<Vec<_>>();
        data.sort_unstable();
        data.dedup();
        data
    }

    /// Drop `object` as an owner of every data item; a data record left
    /// without owners is removed from this component (the data item itself
    /// stays in the package).
    pub(in crate::package) fn remove_data_owner(&mut self, object: u64) {
        for entry in &mut self.data_references {
            entry.owners.retain(|(owner, _)| *owner != object);
        }
        self.data_references
            .retain(|entry| !entry.owners.is_empty());
    }

    /// A copy registered as a new component `identifier` stored at
    /// `Index/<locator>.iwa`, with every registry record keyed by `remap`
    /// (source object → clone). Object UUIDs are not copied: the caller
    /// derives fresh ones. A source with an explicit locator receives the
    /// new explicit locator and keeps its preferred locator unless
    /// `preferred` replaces it, as native Keynote writes slide components
    /// (`Slide` / `Slide-<id>`).
    ///
    /// Versioned external references and ambiguous object identifiers are
    /// not modelled and refuse the clone.
    pub(in crate::package) fn cloned(
        &self,
        identifier: u64,
        locator: &str,
        preferred: Option<&str>,
        remap: &dyn Fn(u64) -> Option<u64>,
    ) -> Result<Self, StructureError> {
        if locator.is_empty() || locator.contains('/') {
            return Err(StructureError::Registry("component locator"));
        }
        let mut explicit = false;
        let mut retained = Vec::with_capacity(self.other.len());
        for raw in &self.other {
            match field_number(raw) {
                Some(COMPONENT_VERSIONED_EXTERNAL_REFERENCES | COMPONENT_AMBIGUOUS_IDENTIFIERS) => {
                    return Err(StructureError::UnsupportedRegistry);
                },
                Some(COMPONENT_IDENTIFIER) => {},
                Some(COMPONENT_LOCATOR) => explicit = true,
                _ => retained.push(raw.clone()),
            }
        }
        let mut other = Vec::with_capacity(retained.len() + 2);
        let mut field = Vec::new();
        encode_varint_into(&mut field, u64::from(COMPONENT_IDENTIFIER << 3));
        encode_varint_into(&mut field, identifier);
        other.push(field);
        for raw in retained {
            if field_number(&raw) == Some(COMPONENT_PREFERRED_LOCATOR) {
                if explicit {
                    other.push(preferred.map_or(raw, |preferred| {
                        string_field(COMPONENT_PREFERRED_LOCATOR, preferred)
                    }));
                    other.push(string_field(COMPONENT_LOCATOR, locator));
                } else {
                    other.push(string_field(COMPONENT_PREFERRED_LOCATOR, locator));
                }
            } else {
                other.push(raw);
            }
        }
        Ok(Self {
            identifier,
            locator: locator.to_owned(),
            external_references: self
                .external_references
                .iter()
                .map(|reference| ExternalReference {
                    object: reference
                        .object
                        .map(|object| remap(object).unwrap_or(object)),
                    ..*reference
                })
                .collect(),
            data_references: self
                .data_references
                .iter()
                .map(|entry| DataReference {
                    data: entry.data,
                    owners: entry
                        .owners
                        .iter()
                        .filter_map(|(owner, count)| remap(*owner).map(|clone| (clone, *count)))
                        .collect(),
                })
                .filter(|entry| !entry.owners.is_empty())
                .collect(),
            object_uuids: Vec::new(),
            save_token: self.save_token,
            other,
        })
    }

    /// Replace the external references.
    pub(in crate::package) fn set_external_references(
        &mut self,
        references: Vec<ExternalReference>,
    ) {
        self.external_references = references;
    }

    fn decode(payload: &[u8], limits: WireLimits) -> Result<Self, StructureError> {
        let mut identifier = None;
        let mut preferred = None;
        let mut explicit = None;
        let mut component = Self {
            identifier: 0,
            locator: String::new(),
            external_references: Vec::new(),
            data_references: Vec::new(),
            object_uuids: Vec::new(),
            save_token: None,
            other: Vec::new(),
        };
        for field in parse_wire_fields(payload, limits)? {
            match (field.number(), field.wire_type()) {
                (COMPONENT_EXTERNAL_REFERENCES, LENGTH_DELIMITED) => {
                    let inner = field.payload(payload)?;
                    let mut reference = ExternalReference {
                        component: 0,
                        object: None,
                        is_weak: None,
                    };
                    let mut has_component = false;
                    for part in parse_wire_fields(inner, limits)? {
                        match (part.number(), part.wire_type()) {
                            (1, VARINT) => {
                                reference.component = varint(part, inner)?;
                                has_component = true;
                            },
                            (2, VARINT) => reference.object = Some(varint(part, inner)?),
                            (3, VARINT) => reference.is_weak = Some(varint(part, inner)? != 0),
                            _ => return Err(StructureError::Registry("external reference field")),
                        }
                    }
                    if !has_component {
                        return Err(StructureError::Registry("external reference component"));
                    }
                    component.external_references.push(reference);
                },
                (COMPONENT_DATA_REFERENCES, LENGTH_DELIMITED) => {
                    let inner = field.payload(payload)?;
                    let mut data = None;
                    let mut owners = Vec::new();
                    for part in parse_wire_fields(inner, limits)? {
                        match (part.number(), part.wire_type()) {
                            (1, VARINT) => data = Some(varint(part, inner)?),
                            (2, LENGTH_DELIMITED) => {
                                let owner = part.payload(inner)?;
                                let mut object = None;
                                let mut count = None;
                                for item in parse_wire_fields(owner, limits)? {
                                    match (item.number(), item.wire_type()) {
                                        (1, VARINT) => object = Some(varint(item, owner)?),
                                        (2, VARINT) => {
                                            count =
                                                Some(u32::try_from(varint(item, owner)?).map_err(
                                                    |_| {
                                                        StructureError::Registry("data owner count")
                                                    },
                                                )?);
                                        },
                                        _ => {
                                            return Err(StructureError::Registry(
                                                "data owner field",
                                            ));
                                        },
                                    }
                                }
                                owners.push((
                                    object.ok_or(StructureError::Registry("data owner object"))?,
                                    count.ok_or(StructureError::Registry("data owner count"))?,
                                ));
                            },
                            _ => return Err(StructureError::Registry("data reference field")),
                        }
                    }
                    component.data_references.push(DataReference {
                        data: data.ok_or(StructureError::Registry("data reference identifier"))?,
                        owners,
                    });
                },
                (COMPONENT_OBJECT_UUIDS, LENGTH_DELIMITED) => {
                    let inner = field.payload(payload)?;
                    let mut object = None;
                    let mut uuid = None;
                    for part in parse_wire_fields(inner, limits)? {
                        match (part.number(), part.wire_type()) {
                            (1, VARINT) => object = Some(varint(part, inner)?),
                            (2, LENGTH_DELIMITED) => {
                                uuid = Some(decode_uuid(part.payload(inner)?, limits)?)
                            },
                            _ => return Err(StructureError::Registry("object UUID field")),
                        }
                    }
                    component.object_uuids.push((
                        object.ok_or(StructureError::Registry("object UUID identifier"))?,
                        uuid.ok_or(StructureError::Registry("object UUID value"))?,
                    ));
                },
                (COMPONENT_SAVE_TOKEN, VARINT) => {
                    if component.save_token.is_some() {
                        return Err(StructureError::Registry("duplicate save token"));
                    }
                    component.save_token = Some(varint(field, payload)?);
                },
                (number, wire_type) => {
                    match (number, wire_type) {
                        (COMPONENT_IDENTIFIER, VARINT) => {
                            if identifier.replace(varint(field, payload)?).is_some() {
                                return Err(StructureError::Registry(
                                    "duplicate component identifier",
                                ));
                            }
                        },
                        (COMPONENT_PREFERRED_LOCATOR, LENGTH_DELIMITED) => {
                            preferred = Some(utf8(field, payload)?);
                        },
                        (COMPONENT_LOCATOR, LENGTH_DELIMITED) => {
                            explicit = Some(utf8(field, payload)?);
                        },
                        (
                            COMPONENT_IDENTIFIER | COMPONENT_PREFERRED_LOCATOR | COMPONENT_LOCATOR,
                            _,
                        ) => {
                            return Err(StructureError::Registry("component header wire type"));
                        },
                        _ => {},
                    }
                    component.other.push(field.raw(payload)?.to_vec());
                },
            }
        }
        component.identifier =
            identifier.ok_or(StructureError::Registry("component identifier"))?;
        component.locator = explicit
            .or(preferred)
            .ok_or(StructureError::Registry("component locator"))?;
        Ok(component)
    }

    fn encode(&self, output: &mut Vec<u8>) {
        for raw in &self.other {
            output.extend_from_slice(raw);
        }
        for reference in &self.external_references {
            let mut inner = Vec::new();
            inner.push(0x08);
            encode_varint_into(&mut inner, reference.component);
            if let Some(object) = reference.object {
                inner.push(0x10);
                encode_varint_into(&mut inner, object);
            }
            if let Some(is_weak) = reference.is_weak {
                inner.push(0x18);
                inner.push(u8::from(is_weak));
            }
            message_field(output, COMPONENT_EXTERNAL_REFERENCES, &inner);
        }
        for entry in &self.data_references {
            let mut inner = Vec::new();
            inner.push(0x08);
            encode_varint_into(&mut inner, entry.data);
            for (object, count) in &entry.owners {
                let mut owner = Vec::new();
                owner.push(0x08);
                encode_varint_into(&mut owner, *object);
                owner.push(0x10);
                encode_varint_into(&mut owner, u64::from(*count));
                message_field(&mut inner, 2, &owner);
            }
            message_field(output, COMPONENT_DATA_REFERENCES, &inner);
        }
        for (object, uuid) in &self.object_uuids {
            let mut inner = Vec::new();
            inner.push(0x08);
            encode_varint_into(&mut inner, *object);
            let mut value = Vec::new();
            value.push(0x08);
            encode_varint_into(&mut value, uuid.lower);
            value.push(0x10);
            encode_varint_into(&mut value, uuid.upper);
            message_field(&mut inner, 2, &value);
            message_field(output, COMPONENT_OBJECT_UUIDS, &inner);
        }
        if let Some(token) = self.save_token {
            encode_varint_into(output, u64::from(COMPONENT_SAVE_TOKEN << 3));
            encode_varint_into(output, token);
        }
    }
}

/// One top-level registry field, in source order.
#[derive(Debug, Clone)]
enum Entry {
    /// An unmodelled field, raw.
    Raw(Vec<u8>),
    /// A component, and its exact source bytes while unchanged.
    Component(Box<Component>, Option<Vec<u8>>),
    /// The object watermark (`last_object_identifier`), and its exact
    /// source bytes while unchanged.
    Watermark(Option<Vec<u8>>),
    /// The root save token, and its exact source bytes while unchanged.
    SaveToken(Option<Vec<u8>>),
}

/// The editable package registry.
#[derive(Debug, Clone)]
pub(in crate::package) struct Registry {
    entries: Vec<Entry>,
    last_object_identifier: u64,
    save_token: Option<u64>,
    touched: Vec<u64>,
}

impl Registry {
    /// Decode a `TSP.PackageMetadata` payload.
    pub(in crate::package) fn decode(
        payload: &[u8],
        limits: WireLimits,
    ) -> Result<Self, StructureError> {
        let mut entries = Vec::new();
        let mut last = None;
        let mut save_token = None;
        for field in parse_wire_fields(payload, limits)? {
            let raw = field.raw(payload)?.to_vec();
            match (field.number(), field.wire_type()) {
                (LAST_OBJECT_IDENTIFIER, VARINT) => {
                    if last.replace(varint(field, payload)?).is_some() {
                        return Err(StructureError::Registry("duplicate object watermark"));
                    }
                    entries.push(Entry::Watermark(Some(raw)));
                },
                (ROOT_SAVE_TOKEN, VARINT) => {
                    if save_token.replace(varint(field, payload)?).is_some() {
                        return Err(StructureError::Registry("duplicate save token"));
                    }
                    entries.push(Entry::SaveToken(Some(raw)));
                },
                (LAST_OBJECT_IDENTIFIER | ROOT_SAVE_TOKEN, _) => {
                    return Err(StructureError::Registry("registry scalar wire type"));
                },
                (COMPONENTS, LENGTH_DELIMITED) => {
                    let component = Component::decode(field.payload(payload)?, limits)?;
                    entries.push(Entry::Component(Box::new(component), Some(raw)));
                },
                _ => entries.push(Entry::Raw(raw)),
            }
        }
        let registry = Self {
            entries,
            last_object_identifier: last.ok_or(StructureError::Registry("object watermark"))?,
            save_token,
            touched: Vec::new(),
        };
        let mut identifiers = registry
            .components()
            .map(|component| component.identifier)
            .collect::<Vec<_>>();
        identifiers.sort_unstable();
        if identifiers.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(StructureError::Registry("duplicate component"));
        }
        Ok(registry)
    }

    pub(in crate::package) const fn last_object_identifier(&self) -> u64 {
        self.last_object_identifier
    }

    /// Raise the object watermark to at least `value`.
    pub(in crate::package) fn set_last_object_identifier(&mut self, value: u64) {
        if value > self.last_object_identifier {
            self.last_object_identifier = value;
            for entry in &mut self.entries {
                if let Entry::Watermark(raw) = entry {
                    *raw = None;
                }
            }
        }
    }

    /// Lower the object watermark to `value` after a removal released the
    /// identifier suffix above it (the caller proves no surviving object
    /// lies above `value`).
    pub(in crate::package) fn release_last_object_identifier(&mut self, value: u64) {
        if value < self.last_object_identifier {
            self.last_object_identifier = value;
            for entry in &mut self.entries {
                if let Entry::Watermark(raw) = entry {
                    *raw = None;
                }
            }
        }
    }

    /// Whether any component still records an owner of data item `data`.
    pub(in crate::package) fn data_has_owner(&self, data: u64) -> bool {
        self.components().any(|component| {
            component
                .data_references
                .iter()
                .any(|entry| entry.data == data)
        })
    }

    /// Remove the `DataInfo` record of `identifier`; returns its stored file
    /// name (`file_name`, else `preferred_file_name`), or `None` when the
    /// registry has no such record.
    pub(in crate::package) fn remove_data_info(
        &mut self,
        identifier: u64,
        limits: WireLimits,
    ) -> Result<Option<String>, StructureError> {
        let mut found = None;
        for (position, entry) in self.entries.iter().enumerate() {
            let Entry::Raw(raw) = entry else { continue };
            if field_number(raw) != Some(DATAS) {
                continue;
            }
            let outer = parse_wire_fields(raw, limits)?;
            let [field] = outer.as_slice() else {
                return Err(StructureError::Registry("data info"));
            };
            let inner = field.payload(raw)?;
            let mut id = None;
            let mut preferred = None;
            let mut stored = None;
            for part in parse_wire_fields(inner, limits)? {
                match (part.number(), part.wire_type()) {
                    (1, VARINT) => id = Some(varint(part, inner)?),
                    (3, LENGTH_DELIMITED) => preferred = Some(part.payload(inner)?.to_vec()),
                    (4, LENGTH_DELIMITED) => stored = Some(part.payload(inner)?.to_vec()),
                    _ => {},
                }
            }
            if id == Some(identifier) {
                if found.is_some() {
                    return Err(StructureError::Registry("duplicate data info"));
                }
                let name = stored
                    .or(preferred)
                    .map(String::from_utf8)
                    .transpose()
                    .map_err(|_| StructureError::Registry("data file name"))?
                    .ok_or(StructureError::Registry("data file name"))?;
                found = Some((position, name));
            }
        }
        Ok(found.map(|(position, name)| {
            self.entries.remove(position);
            name
        }))
    }

    /// The stored file name (`file_name`, else `preferred_file_name`) of the
    /// `DataInfo` record of `identifier`.
    pub(in crate::package) fn data_file_name(
        &self,
        identifier: u64,
        limits: WireLimits,
    ) -> Result<Option<String>, StructureError> {
        let mut found = None;
        for entry in &self.entries {
            let Entry::Raw(raw) = entry else { continue };
            if field_number(raw) != Some(DATAS) {
                continue;
            }
            let outer = parse_wire_fields(raw, limits)?;
            let [field] = outer.as_slice() else {
                return Err(StructureError::Registry("data info"));
            };
            let inner = field.payload(raw)?;
            let mut id = None;
            let mut preferred = None;
            let mut stored = None;
            for part in parse_wire_fields(inner, limits)? {
                match (part.number(), part.wire_type()) {
                    (1, VARINT) => id = Some(varint(part, inner)?),
                    (3, LENGTH_DELIMITED) => preferred = Some(part.payload(inner)?),
                    (4, LENGTH_DELIMITED) => stored = Some(part.payload(inner)?),
                    _ => {},
                }
            }
            if id == Some(identifier) {
                let name = std::str::from_utf8(
                    stored
                        .or(preferred)
                        .ok_or(StructureError::Registry("data file name"))?,
                )
                .map_err(|_| StructureError::Registry("data file name"))?;
                if found.replace(name.to_owned()).is_some() {
                    return Err(StructureError::Registry("duplicate data info"));
                }
            }
        }
        Ok(found)
    }

    /// Replace the SHA-1 `digest` (field 2) and materialized length (field
    /// 18) of the `DataInfo` record of `identifier`, keeping every other
    /// field and the record's position; returns the record's stored file
    /// name, or `None` when the registry has no such record.
    pub(in crate::package) fn patch_data_info(
        &mut self,
        identifier: u64,
        digest: &[u8],
        length: u64,
        limits: WireLimits,
    ) -> Result<Option<String>, StructureError> {
        let mut found = None;
        for (position, entry) in self.entries.iter().enumerate() {
            let Entry::Raw(raw) = entry else { continue };
            if field_number(raw) != Some(DATAS) {
                continue;
            }
            let outer = parse_wire_fields(raw, limits)?;
            let [field] = outer.as_slice() else {
                return Err(StructureError::Registry("data info"));
            };
            let inner = field.payload(raw)?;
            let mut id = None;
            let mut preferred = None;
            let mut stored = None;
            let mut patched = Vec::with_capacity(inner.len() + digest.len());
            let (mut digests, mut lengths) = (0, 0);
            for part in parse_wire_fields(inner, limits)? {
                match (part.number(), part.wire_type()) {
                    (1, VARINT) => id = Some(varint(part, inner)?),
                    (3, LENGTH_DELIMITED) => preferred = Some(part.payload(inner)?.to_vec()),
                    (4, LENGTH_DELIMITED) => stored = Some(part.payload(inner)?.to_vec()),
                    _ => {},
                }
                match (part.number(), part.wire_type()) {
                    (2, LENGTH_DELIMITED) => {
                        digests += 1;
                        message_field(&mut patched, 2, digest);
                    },
                    (18, VARINT) => {
                        lengths += 1;
                        encode_varint_into(&mut patched, u64::from(18_u32 << 3));
                        encode_varint_into(&mut patched, length);
                    },
                    _ => patched.extend_from_slice(part.raw(inner)?),
                }
            }
            if id != Some(identifier) {
                continue;
            }
            if found.is_some() || digests > 1 || lengths > 1 {
                return Err(StructureError::Registry("data info digest"));
            }
            if digests == 0 {
                message_field(&mut patched, 2, digest);
            }
            if lengths == 0 {
                encode_varint_into(&mut patched, u64::from(18_u32 << 3));
                encode_varint_into(&mut patched, length);
            }
            let name = stored
                .or(preferred)
                .map(String::from_utf8)
                .transpose()
                .map_err(|_| StructureError::Registry("data file name"))?
                .ok_or(StructureError::Registry("data file name"))?;
            let mut replacement = Vec::with_capacity(patched.len() + 4);
            message_field(&mut replacement, DATAS, &patched);
            found = Some((position, name, replacement));
        }
        Ok(found.map(|(position, name, replacement)| {
            self.entries[position] = Entry::Raw(replacement);
            name
        }))
    }

    pub(in crate::package) fn components(&self) -> impl Iterator<Item = &Component> {
        self.entries.iter().filter_map(|entry| match entry {
            Entry::Component(component, _) => Some(component.as_ref()),
            _ => None,
        })
    }

    /// The component registered for a package member name.
    pub(in crate::package) fn component_for_member(&self, member: &str) -> Option<&Component> {
        self.components()
            .find(|component| component.member_name() == member)
    }

    pub(in crate::package) fn component(&self, identifier: u64) -> Option<&Component> {
        self.components()
            .find(|component| component.identifier == identifier)
    }

    /// Mutably borrow a component, marking it touched (its save token
    /// advances on [`Self::encode`]).
    pub(in crate::package) fn component_mut(&mut self, identifier: u64) -> Option<&mut Component> {
        let found = self.entries.iter_mut().find_map(|entry| match entry {
            Entry::Component(component, raw) if component.identifier == identifier => {
                *raw = None;
                Some(component.as_mut())
            },
            _ => None,
        })?;
        if !self.touched.contains(&identifier) {
            self.touched.push(identifier);
        }
        Some(found)
    }

    /// Mutably borrow a component without advancing its save token, for a
    /// registry-only preparation whose following transaction rewrites the
    /// same component and advances the token once.
    pub(in crate::package) fn component_mut_untracked(
        &mut self,
        identifier: u64,
    ) -> Option<&mut Component> {
        self.entries.iter_mut().find_map(|entry| match entry {
            Entry::Component(component, raw) if component.identifier == identifier => {
                *raw = None;
                Some(component.as_mut())
            },
            _ => None,
        })
    }

    /// Append a new component after the last existing one.
    pub(in crate::package) fn add_component(
        &mut self,
        component: Component,
    ) -> Result<(), StructureError> {
        if component.identifier == 0
            || self.component(component.identifier).is_some()
            || self
                .components()
                .any(|existing| existing.locator == component.locator)
        {
            return Err(StructureError::Registry("component already registered"));
        }
        self.touched.push(component.identifier);
        let position = self
            .entries
            .iter()
            .rposition(|entry| matches!(entry, Entry::Component(..)))
            .map_or(self.entries.len(), |index| index + 1);
        self.entries
            .insert(position, Entry::Component(Box::new(component), None));
        Ok(())
    }

    /// Every registered data identifier (`DataInfo.identifier`).
    pub(in crate::package) fn data_identifiers(
        &self,
        limits: WireLimits,
    ) -> Result<Vec<u64>, StructureError> {
        let mut identifiers = Vec::new();
        for entry in &self.entries {
            let Entry::Raw(raw) = entry else { continue };
            if field_number(raw) != Some(DATAS) {
                continue;
            }
            let outer = parse_wire_fields(raw, limits)?;
            let [field] = outer.as_slice() else {
                return Err(StructureError::Registry("data info"));
            };
            let inner = field.payload(raw)?;
            let mut identifier = None;
            for part in parse_wire_fields(inner, limits)? {
                if part.number() == 1 && part.wire_type() == VARINT {
                    identifier = Some(varint(part, inner)?);
                }
            }
            identifiers.push(identifier.ok_or(StructureError::Registry("data identifier"))?);
        }
        Ok(identifiers)
    }

    /// Append one `DataInfo` record (identifier, SHA-1 digest, preferred and
    /// stored file names and materialized length) after the existing ones.
    pub(in crate::package) fn add_data_info(
        &mut self,
        identifier: u64,
        digest: &[u8],
        preferred_file_name: &str,
        file_name: &str,
        length: u64,
    ) {
        let mut inner = Vec::new();
        inner.push(0x08);
        encode_varint_into(&mut inner, identifier);
        message_field(&mut inner, 2, digest);
        message_field(&mut inner, 3, preferred_file_name.as_bytes());
        message_field(&mut inner, 4, file_name.as_bytes());
        encode_varint_into(&mut inner, u64::from(18_u32 << 3));
        encode_varint_into(&mut inner, length);
        let mut raw = Vec::new();
        message_field(&mut raw, DATAS, &inner);
        let position = self
            .entries
            .iter()
            .rposition(|entry| match entry {
                Entry::Raw(raw) => field_number(raw) == Some(DATAS),
                Entry::Component(..) => true,
                _ => false,
            })
            .map_or(self.entries.len(), |index| index + 1);
        self.entries.insert(position, Entry::Raw(raw));
    }

    /// Encode the registry. Unchanged fields keep their exact bytes and
    /// order; when any component was touched the root save token advances
    /// once and is stamped on every touched component.
    pub(in crate::package) fn encode(&self) -> Result<Vec<u8>, StructureError> {
        let advanced = if self.touched.is_empty() {
            None
        } else {
            Some(
                self.save_token
                    .unwrap_or(0)
                    .checked_add(1)
                    .ok_or(StructureError::Registry("save token overflow"))?,
            )
        };
        let mut output = Vec::new();
        let mut wrote_token = false;
        for entry in &self.entries {
            match entry {
                Entry::Raw(raw) | Entry::Component(_, Some(raw)) | Entry::Watermark(Some(raw)) => {
                    output.extend_from_slice(raw)
                },
                Entry::Watermark(None) => {
                    encode_varint_into(&mut output, u64::from(LAST_OBJECT_IDENTIFIER << 3));
                    encode_varint_into(&mut output, self.last_object_identifier);
                },
                Entry::SaveToken(raw) => {
                    wrote_token = true;
                    match (advanced, raw) {
                        (None, Some(raw)) => output.extend_from_slice(raw),
                        (token, _) => {
                            encode_varint_into(&mut output, u64::from(ROOT_SAVE_TOKEN << 3));
                            encode_varint_into(&mut output, token.or(self.save_token).unwrap_or(0));
                        },
                    }
                },
                Entry::Component(component, None) => {
                    let mut component = component.as_ref().clone();
                    if let Some(token) = advanced
                        && self.touched.contains(&component.identifier)
                    {
                        component.save_token = Some(token);
                    }
                    let mut inner = Vec::new();
                    component.encode(&mut inner);
                    message_field(&mut output, COMPONENTS, &inner);
                },
            }
        }
        if let (Some(token), false) = (advanced, wrote_token) {
            encode_varint_into(&mut output, u64::from(ROOT_SAVE_TOKEN << 3));
            encode_varint_into(&mut output, token);
        }
        Ok(output)
    }
}

fn field_number(raw: &[u8]) -> Option<u32> {
    decode_varint_from_bytes(raw)
        .ok()
        .and_then(|(key, _)| u32::try_from(key >> 3).ok())
}

fn varint(field: WireField, payload: &[u8]) -> Result<u64, StructureError> {
    let bytes = field.payload(payload)?;
    let (value, used) =
        decode_varint_from_bytes(bytes).map_err(|_| StructureError::Registry("varint"))?;
    if used != bytes.len() {
        return Err(StructureError::Registry("varint"));
    }
    Ok(value)
}

fn utf8(field: WireField, payload: &[u8]) -> Result<String, StructureError> {
    std::str::from_utf8(field.payload(payload)?)
        .map(str::to_owned)
        .map_err(|_| StructureError::Registry("locator UTF-8"))
}

fn decode_uuid(payload: &[u8], limits: WireLimits) -> Result<Uuid, StructureError> {
    let mut lower = None;
    let mut upper = None;
    for field in parse_wire_fields(payload, limits)? {
        match (field.number(), field.wire_type()) {
            (1, VARINT) => lower = Some(varint(field, payload)?),
            (2, VARINT) => upper = Some(varint(field, payload)?),
            _ => return Err(StructureError::Registry("UUID field")),
        }
    }
    Ok(Uuid {
        lower: lower.ok_or(StructureError::Registry("UUID lower"))?,
        upper: upper.ok_or(StructureError::Registry("UUID upper"))?,
    })
}

fn string_field(number: u32, value: &str) -> Vec<u8> {
    let mut output = Vec::new();
    message_field(&mut output, number, value.as_bytes());
    output
}

pub(in crate::package) fn message_field(output: &mut Vec<u8>, number: u32, inner: &[u8]) {
    encode_varint_into(output, u64::from(number << 3 | 2));
    encode_varint_into(output, inner.len() as u64);
    output.extend_from_slice(inner);
}

#[cfg(test)]
mod tests {
    use super::super::test_support::{corpus, fixture, registry_payload};
    use super::*;

    fn limits() -> WireLimits {
        WireLimits::default()
    }

    fn raw_components(payload: &[u8]) -> Vec<Vec<u8>> {
        parse_wire_fields(payload, limits())
            .expect("registry fields")
            .into_iter()
            .filter(|field| field.number() == COMPONENTS)
            .map(|field| field.raw(payload).expect("raw component").to_vec())
            .collect()
    }

    #[test]
    fn untouched_registries_reencode_exactly() {
        for (path, package) in corpus() {
            let payload = registry_payload(&package);
            let registry = Registry::decode(&payload, limits()).expect("registry decodes");
            assert_eq!(registry.encode().expect("encodes"), payload, "{path:?}");
            let again = Registry::decode(&registry.encode().expect("encodes"), limits())
                .expect("re-decodes");
            assert_eq!(
                again.components().count(),
                registry.components().count(),
                "{path:?}"
            );
        }
    }

    #[test]
    fn touched_components_advance_save_tokens_and_keep_the_rest() {
        let payload = registry_payload(&fixture("basic.key"));
        let mut registry = Registry::decode(&payload, limits()).expect("registry decodes");
        let token = registry.save_token.expect("root save token");
        let slide = registry
            .component_for_member("Index/Slide-2652150.iwa")
            .expect("slide component")
            .identifier;
        let before = registry.component(slide).expect("slide").clone();
        registry
            .component_mut(slide)
            .expect("slide")
            .add_object_uuid(9_999_999, Uuid { lower: 7, upper: 9 });
        registry.set_last_object_identifier(1);
        assert_eq!(
            registry.last_object_identifier(),
            2_652_562,
            "watermark only rises"
        );
        registry.set_last_object_identifier(2_700_000);
        let encoded = registry.encode().expect("encodes");
        let decoded = Registry::decode(&encoded, limits()).expect("re-decodes");
        assert_eq!(decoded.save_token, Some(token + 1));
        assert_eq!(decoded.last_object_identifier(), 2_700_000);
        let after = decoded.component(slide).expect("slide");
        assert_eq!(after.save_token, Some(token + 1));
        assert_eq!(after.object_uuids().len(), before.object_uuids().len() + 1);
        assert_eq!(after.uuid_of(9_999_999), Some(Uuid { lower: 7, upper: 9 }));
        assert_eq!(after.external_references(), before.external_references());
        let source = raw_components(&payload);
        let target = raw_components(&encoded);
        assert_eq!(source.len(), target.len());
        let changed = source
            .iter()
            .zip(&target)
            .filter(|(source, target)| source != target)
            .count();
        assert_eq!(changed, 1, "only the touched component is re-encoded");
    }

    #[test]
    fn cloned_components_take_slide_locators_and_remapped_records() {
        let payload = registry_payload(&fixture("basic.key"));
        let mut registry = Registry::decode(&payload, limits()).expect("registry decodes");
        let slide = registry
            .component_for_member("Index/Slide-2652150.iwa")
            .expect("slide component")
            .clone();
        let clone = slide
            .cloned(42, "Slide-42", None, &|object| {
                (object == 2_652_150).then_some(42)
            })
            .expect("clones");
        assert_eq!(clone.identifier, 42);
        assert_eq!(clone.member_name(), "Index/Slide-42.iwa");
        assert!(clone.object_uuids().is_empty());
        assert_eq!(clone.external_references(), slide.external_references());
        let renamed = slide
            .cloned(43, "Slide-43", Some("Slide"), &|_| None)
            .expect("clones");
        let mut encoded = Vec::new();
        renamed.encode(&mut encoded);
        let decoded = Component::decode(&encoded, limits()).expect("decodes");
        assert_eq!(decoded.locator, "Slide-43");
        assert_eq!(decoded.identifier, 43);
        assert!(slide.cloned(44, "a/b", None, &|_| None).is_err());
        registry.add_component(clone).expect("registers");
        let duplicate = slide
            .cloned(42, "Slide-other", None, &|_| None)
            .expect("clones");
        assert!(registry.add_component(duplicate).is_err());
        let reencoded = registry.encode().expect("encodes");
        let decoded = Registry::decode(&reencoded, limits()).expect("re-decodes");
        assert!(decoded.component_for_member("Index/Slide-42.iwa").is_some());
    }

    #[test]
    fn versioned_records_refuse_cloning() {
        let mut component = Component::decode(
            &{
                let mut output = Vec::new();
                output.extend_from_slice(&[0x08, 0x05]);
                message_field(&mut output, COMPONENT_PREFERRED_LOCATOR, b"Slide");
                message_field(
                    &mut output,
                    COMPONENT_VERSIONED_EXTERNAL_REFERENCES,
                    &[0x08, 0x01],
                );
                output
            },
            limits(),
        )
        .expect("decodes");
        assert_eq!(
            component.cloned(6, "Slide-6", None, &|_| None).err(),
            Some(StructureError::UnsupportedRegistry)
        );
        component.other.pop();
        assert!(component.cloned(6, "Slide-6", None, &|_| None).is_ok());
    }

    #[test]
    fn data_owners_are_counted() {
        let mut component = Component::decode(
            &{
                let mut output = vec![0x08, 0x05];
                message_field(&mut output, COMPONENT_PREFERRED_LOCATOR, b"Slide");
                output
            },
            limits(),
        )
        .expect("decodes");
        component.add_data_owner(9, 100);
        component.add_data_owner(9, 100);
        component.add_data_owner(9, 101);
        component.add_external_reference(ExternalReference {
            component: 1,
            object: None,
            is_weak: None,
        });
        component.add_external_reference(ExternalReference {
            component: 1,
            object: None,
            is_weak: None,
        });
        let mut encoded = Vec::new();
        component.encode(&mut encoded);
        let decoded = Component::decode(&encoded, limits()).expect("decodes");
        assert_eq!(
            decoded.data_references,
            vec![DataReference {
                data: 9,
                owners: vec![(100, 2), (101, 1)],
            }]
        );
        assert_eq!(decoded.external_references().len(), 1);
    }

    #[test]
    fn data_owners_and_uuids_are_removed() {
        let mut component = Component::decode(
            &{
                let mut output = vec![0x08, 0x05];
                message_field(&mut output, COMPONENT_PREFERRED_LOCATOR, b"Slide");
                output
            },
            limits(),
        )
        .expect("decodes");
        component.add_data_owner(9, 100);
        component.add_data_owner(9, 101);
        component.add_data_owner(10, 100);
        component.add_object_uuid(100, Uuid { lower: 1, upper: 2 });
        component.add_object_uuid(101, Uuid { lower: 3, upper: 4 });
        component.remove_data_owner(100);
        assert_eq!(
            component.data_references,
            vec![DataReference {
                data: 9,
                owners: vec![(101, 1)],
            }]
        );
        assert_eq!(component.remove_object_uuids(&[100, 7]), 1);
        assert_eq!(component.uuid_of(100), None);
        assert_eq!(component.uuid_of(101), Some(Uuid { lower: 3, upper: 4 }));
    }

    #[test]
    fn data_infos_are_listed_and_appended() {
        let package = fixture("basic.key");
        let mut registry =
            Registry::decode(&registry_payload(&package), limits()).expect("decodes");
        let before = registry.data_identifiers(limits()).expect("lists");
        assert!(!before.is_empty());
        let next = before.iter().max().copied().unwrap_or(0) + 1;
        registry.add_data_info(next, &[7; 20], "pixel.png", "pixel-1.png", 74);
        let encoded = registry.encode().expect("encodes");
        let decoded = Registry::decode(&encoded, limits()).expect("decodes");
        let after = decoded.data_identifiers(limits()).expect("lists");
        assert_eq!(after.len(), before.len() + 1);
        assert_eq!(after.last(), Some(&next));
    }
}
