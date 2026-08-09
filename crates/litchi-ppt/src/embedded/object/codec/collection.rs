//! Ordered `ExObjList` parsing and lossless unknown-record serialization.

use super::super::model::{
    Collection, Control, Definition, ExternalObject, MAX_OLE_OBJECTS, UnknownRecord,
};
use super::wire::{corrupted, record_bytes, record_bytes_raw, require_atom};
use crate::consts::RecordType;
use crate::external_media::Collection as MediaCollection;
use crate::hyperlink::Hyperlinks;
use crate::package::{Error, Result};
use crate::records::Record;
use std::collections::HashSet;

impl UnknownRecord {
    #[must_use]
    pub fn record_type(&self) -> u16 {
        self.record.record_type_raw
    }

    #[must_use]
    pub fn version(&self) -> u16 {
        self.record.version
    }

    #[must_use]
    pub fn instance(&self) -> u16 {
        self.record.instance
    }

    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.record.data
    }

    /// Serialize to the raw bytes of the preserved unknown record.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header fields or the payload exceed
    /// the encodable domain.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        record_bytes_raw(
            self.record.version,
            self.record.instance,
            self.record.record_type_raw,
            &self.record.data,
        )
    }

    pub(crate) fn from_record(record: &Record, object_index: usize) -> Self {
        Self {
            record: record.clone(),
            object_index,
        }
    }
}

impl Collection {
    /// Parse the `ExObjList` tree of a `PowerPoint` document record.
    ///
    /// Returns `Ok(None)` when the record tree contains no external-object
    /// list.
    ///
    /// # Errors
    ///
    /// Returns an error if the record tree contains multiple external-object
    /// lists, if any record header, size, or fixed-width field is invalid, or
    /// if object IDs violate the seed, uniqueness, or cross-collection
    /// invariants.
    pub fn parse(root: &Record) -> Result<Option<Self>> {
        let mut lists = Vec::new();
        collect_external_object_lists(root, &mut lists);
        if lists.len() > 1 {
            return corrupted("record tree contains multiple external-object lists");
        }
        let Some(list) = lists.first() else {
            return Ok(None);
        };
        if list.version != 0x0f || list.instance != 0 || list.record_type != RecordType::ExObjList {
            return corrupted("ExObjListContainer has an invalid header");
        }
        let children = Record::parse_sequence_strict(&list.data, "ExObjListContainer")?;
        let Some(atom) = children.first() else {
            return corrupted("ExObjListContainer is missing ExObjListAtom");
        };
        require_atom(atom, 0, 0, RecordType::ExObjListAtom, 4, "ExObjListAtom")?;
        let signed_seed =
            i32::from_le_bytes([atom.data[0], atom.data[1], atom.data[2], atom.data[3]]);
        if signed_seed < 1 {
            return corrupted("ExObjListAtom identifier seed must be positive");
        }
        let id_seed = signed_seed.cast_unsigned();
        let mut ids = HashSet::new();
        let mut objects = Vec::new();
        let mut unknown_records = Vec::new();
        for child in &children[1..] {
            if !matches!(
                child.record_type,
                RecordType::ExternalOleEmbed
                    | RecordType::ExternalOleLink
                    | RecordType::ExternalOleControl
            ) {
                unknown_records.push(UnknownRecord::from_record(child, objects.len()));
                continue;
            }
            if objects.len() >= MAX_OLE_OBJECTS {
                return corrupted(format!(
                    "external-object list exceeds {MAX_OLE_OBJECTS} OLE objects"
                ));
            }
            let object = if child.record_type == RecordType::ExternalOleControl {
                ExternalObject::ActiveXControl(Control::parse(child)?)
            } else {
                ExternalObject::Object(Definition::parse(child)?)
            };
            let id = object.id();
            if id > id_seed {
                return corrupted(format!(
                    "OLE object ID {id} exceeds ExObjList seed {id_seed}"
                ));
            }
            if !ids.insert(id) {
                return corrupted(format!(
                    "external-object list contains duplicate OLE object ID {id}"
                ));
            }
            objects.push(object);
        }
        if let Some(media) = MediaCollection::parse(root)?
            && objects
                .iter()
                .any(|object| media.get(object.id()).is_some())
        {
            return corrupted("external-object list reuses an ID for OLE and media objects");
        }
        let hyperlinks = Hyperlinks::parse(root)?;
        if objects.iter().any(|object| {
            hyperlinks
                .hyperlinks
                .iter()
                .any(|hyperlink| hyperlink.id == object.id())
        }) {
            return corrupted("external-object list reuses an ID for OLE objects and hyperlinks");
        }
        Ok(Some(Self {
            id_seed,
            objects,
            unknown_records,
        }))
    }

    #[must_use]
    pub fn get(&self, id: u32) -> Option<&ExternalObject> {
        self.objects.iter().find(|object| object.id() == id)
    }

    #[must_use]
    pub fn find(&self, id: u32) -> Option<&ExternalObject> {
        self.get(id)
    }

    #[must_use]
    pub fn unknown_records(&self) -> &[UnknownRecord] {
        &self.unknown_records
    }

    /// Serialize the collection to the raw bytes of an `ExObjList` container,
    /// interleaving preserved unknown records at their original positions.
    ///
    /// # Errors
    ///
    /// Returns an error if the collection violates its validation invariants,
    /// if the identifier seed exceeds the signed-integer domain, or if any
    /// child record fails to serialize.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let seed = i32::try_from(self.id_seed)
            .map_err(|_err| Error::Corrupted("ExObjList identifier seed exceeds i32".into()))?;
        let mut children = record_bytes(0, 0, RecordType::ExObjListAtom, &seed.to_le_bytes())?;
        for object_index in 0..=self.objects.len() {
            for record in self
                .unknown_records
                .iter()
                .filter(|record| record.object_index == object_index)
            {
                children.extend_from_slice(&record.to_record_bytes()?);
            }
            if let Some(object) = self.objects.get(object_index) {
                children.extend_from_slice(&object.to_record_bytes()?);
            }
        }
        record_bytes(0x0f, 0, RecordType::ExObjList, &children)
    }
}

fn collect_external_object_lists<'a>(record: &'a Record, lists: &mut Vec<&'a Record>) {
    if record.record_type == RecordType::ExObjList {
        lists.push(record);
    }
    for child in &record.children {
        collect_external_object_lists(child, lists);
    }
}
