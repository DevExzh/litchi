//! Ordered `ExObjList` parsing and lossless unknown-record serialization.

use super::super::model::{
    Collection, Control, Definition, ExternalObject, MAX_OLE_OBJECTS, UnknownRecord,
};
use super::wire::{
    corrupted, ensure_output_total, ensure_record_size, record_bytes, record_bytes_raw_with_limits,
    record_bytes_with_limits, require_atom,
};
use crate::consts::RecordType;
use crate::external_media::{Collection as MediaCollection, Limits as MediaLimits};
use crate::hyperlink::Hyperlinks;
use crate::package::{Error, Result, SourceLimits};
use crate::records::{Record, RecordParseSession};
use std::borrow::Cow;
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
    pub fn to_record_bytes(&self, limits: &SourceLimits) -> Result<Vec<u8>> {
        record_bytes_raw_with_limits(
            self.record.version,
            self.record.instance,
            self.record.record_type_raw,
            &self.record.data,
            limits,
            "unknown ExObjList record",
        )
    }

    pub(crate) fn from_record(record: &Record, object_index: usize) -> Self {
        Self {
            record: record.clone(),
            object_index,
        }
    }

    pub(crate) fn record_len(&self, limits: &SourceLimits) -> Result<usize> {
        ensure_record_size(self.record.data.len(), limits, "unknown ExObjList record")?;
        self.record
            .data
            .len()
            .checked_add(8)
            .ok_or_else(|| Error::Corrupted("unknown ExObjList record length overflows".into()))
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
    pub fn parse(root: &Record, limits: &SourceLimits) -> Result<Option<Self>> {
        let mut session = RecordParseSession::new(*limits, root.data.len())?;
        let mut lists = Vec::new();
        collect_external_object_lists(root, &mut lists, &mut session)?;
        if lists.len() > 1 {
            return corrupted("record tree contains multiple external-object lists");
        }
        let Some((list, list_depth)) = lists.first().copied() else {
            return Ok(None);
        };
        if list.version != 0x0f || list.instance != 0 || list.record_type != RecordType::ExObjList {
            return corrupted("ExObjListContainer has an invalid header");
        }
        ensure_record_size(list.data.len(), limits, "ExObjListContainer")?;
        let children: Cow<'_, [Record]> = if !list.children.is_empty() || list.data.is_empty() {
            Cow::Borrowed(&list.children)
        } else {
            let child_depth = list_depth
                .checked_add(1)
                .ok_or_else(|| Error::Corrupted("external-object list depth overflows".into()))?;
            Cow::Owned(session.parse_sequence(&list.data, "ExObjListContainer", child_depth)?)
        };
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
        let object_depth = list_depth
            .checked_add(1)
            .ok_or_else(|| Error::Corrupted("external-object depth overflows".into()))?;
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
                ExternalObject::ActiveXControl(Control::parse_with_session(
                    child,
                    limits,
                    &mut session,
                    object_depth,
                )?)
            } else {
                ExternalObject::Object(Definition::parse_with_session(
                    child,
                    limits,
                    &mut session,
                    object_depth,
                )?)
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
        let media_limits = media_limits(*limits);
        if let Some(media) = MediaCollection::parse(root, &media_limits)?
            && objects
                .iter()
                .any(|object| media.get(object.id()).is_some())
        {
            return corrupted("external-object list reuses an ID for OLE and media objects");
        }
        let hyperlinks = Hyperlinks::parse(root, limits)?;
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
    pub fn to_record_bytes(&self, limits: &SourceLimits) -> Result<Vec<u8>> {
        self.validate()?;
        ensure_record_size(4, limits, "ExObjListAtom")?;
        let seed = i32::try_from(self.id_seed)
            .map_err(|_err| Error::Corrupted("ExObjList identifier seed exceeds i32".into()))?;
        let mut children = record_bytes(0, 0, RecordType::ExObjListAtom, &seed.to_le_bytes())?;
        ensure_output_total(children.len(), limits, "ExObjListContainer")?;
        for object_index in 0..=self.objects.len() {
            for record in self
                .unknown_records
                .iter()
                .filter(|record| record.object_index == object_index)
            {
                let child_len = record.record_len(limits)?;
                let length = children.len().checked_add(child_len).ok_or_else(|| {
                    Error::Corrupted("ExObjListContainer length overflows".into())
                })?;
                ensure_record_size(length, limits, "ExObjListContainer")?;
                ensure_output_total(length, limits, "ExObjListContainer")?;
                let bytes = record.to_record_bytes(limits)?;
                debug_assert_eq!(bytes.len(), child_len);
                children.extend_from_slice(&bytes);
            }
            if let Some(object) = self.objects.get(object_index) {
                let child_len = object.record_len(limits)?;
                let length = children.len().checked_add(child_len).ok_or_else(|| {
                    Error::Corrupted("ExObjListContainer length overflows".into())
                })?;
                ensure_record_size(length, limits, "ExObjListContainer")?;
                ensure_output_total(length, limits, "ExObjListContainer")?;
                let bytes = object.to_record_bytes(limits)?;
                debug_assert_eq!(bytes.len(), child_len);
                children.extend_from_slice(&bytes);
            }
        }
        record_bytes_with_limits(
            0x0f,
            0,
            RecordType::ExObjList,
            &children,
            limits,
            "ExObjListContainer",
        )
    }
}

impl ExternalObject {
    pub(crate) fn record_len(&self, limits: &SourceLimits) -> Result<usize> {
        match self {
            Self::Object(value) => value.record_len(limits),
            Self::ActiveXControl(value) => value.record_len(limits),
        }
    }
}

fn media_limits(source: SourceLimits) -> MediaLimits {
    MediaLimits {
        max_root_bytes: source.max_record_bytes().min(source.max_input_bytes()),
        max_record_bytes: source.max_record_bytes(),
        max_depth: source.max_depth(),
        max_records: source.max_events(),
        max_owner_references: source.max_events(),
    }
}

fn collect_external_object_lists<'a>(
    root: &'a Record,
    lists: &mut Vec<(&'a Record, usize)>,
    session: &mut RecordParseSession,
) -> Result<()> {
    let mut pending = Vec::new();
    pending
        .try_reserve(1)
        .map_err(|_error| Error::AllocationFailed("external-object traversal stack"))?;
    pending.push((root, 0usize));
    while let Some((record, depth)) = pending.pop() {
        session.account_existing(record, depth)?;
        if record.record_type == RecordType::ExObjList {
            lists
                .try_reserve(1)
                .map_err(|_error| Error::AllocationFailed("external-object list table"))?;
            lists.push((record, depth));
        }
        let child_depth = depth
            .checked_add(1)
            .ok_or_else(|| Error::Corrupted("external-object traversal depth overflows".into()))?;
        pending
            .try_reserve(record.children.len())
            .map_err(|_error| Error::AllocationFailed("external-object traversal stack"))?;
        pending.extend(record.children.iter().map(|child| (child, child_depth)));
    }
    Ok(())
}
