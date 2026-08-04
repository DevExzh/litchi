//! Metadata-preserving IWA archive parsing, mutation, and serialization.
//!
//! A decompressed IWA stream is a sequence of objects. Each object starts with
//! a varint-sized `TSP.ArchiveInfo` protobuf followed by the payloads described
//! by its `TSP.MessageInfo` entries.

use std::collections::HashSet;
use std::io::{Cursor, Read};

use prost::Message;

use crate::protobuf::{self, DecodedMessage, decode};
use crate::varint;
use crate::{Error, Result};

/// Archive metadata for one object in an IWA component.
#[derive(Debug, Clone, PartialEq)]
pub struct ArchiveInfo {
    /// Unique identifier across the document package.
    pub identifier: Option<u64>,
    /// Metadata for each payload immediately following this header.
    pub message_infos: Vec<MessageInfo>,
    /// Whether this archive should be merged with an existing object.
    pub should_merge: Option<bool>,
}

impl ArchiveInfo {
    /// Decode an `ArchiveInfo` protobuf from a bounded reader.
    pub fn parse<R: Read>(reader: &mut R) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        let message = protobuf::tsp::ArchiveInfo::decode(data.as_slice())?;
        Ok(message.into())
    }

    fn encode(&self) -> Vec<u8> {
        protobuf::tsp::ArchiveInfo::from(self).encode_to_vec()
    }
}

impl From<protobuf::tsp::ArchiveInfo> for ArchiveInfo {
    fn from(value: protobuf::tsp::ArchiveInfo) -> Self {
        Self {
            identifier: value.identifier,
            message_infos: value.message_infos.into_iter().map(Into::into).collect(),
            should_merge: value.should_merge,
        }
    }
}

impl From<&ArchiveInfo> for protobuf::tsp::ArchiveInfo {
    fn from(value: &ArchiveInfo) -> Self {
        Self {
            identifier: value.identifier,
            message_infos: value.message_infos.iter().map(Into::into).collect(),
            should_merge: value.should_merge,
        }
    }
}

/// Metadata for one protobuf payload in an archive object.
#[derive(Debug, Clone, PartialEq)]
pub struct MessageInfo {
    pub type_: u32,
    pub versions: Vec<u32>,
    pub length: u32,
    pub field_infos: Vec<protobuf::tsp::FieldInfo>,
    pub object_references: Vec<u64>,
    pub data_references: Vec<u64>,
    pub base_message_index: Option<u32>,
    pub diff_merge_version: Vec<u32>,
    pub diff_field_path: Option<protobuf::tsp::FieldPath>,
    pub fields_to_remove: Vec<protobuf::tsp::FieldPath>,
    pub diff_read_version: Vec<u32>,
}

impl MessageInfo {
    /// Construct metadata for a new payload.
    pub fn new(type_: u32, length: u32) -> Self {
        Self {
            type_,
            versions: vec![1, 0, 5],
            length,
            field_infos: Vec::new(),
            object_references: Vec::new(),
            data_references: Vec::new(),
            base_message_index: None,
            diff_merge_version: Vec::new(),
            diff_field_path: None,
            fields_to_remove: Vec::new(),
            diff_read_version: Vec::new(),
        }
    }

    /// Decode a `MessageInfo` protobuf from a bounded reader.
    pub fn parse<R: Read>(reader: &mut R) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        let message = protobuf::tsp::MessageInfo::decode(data.as_slice())?;
        Ok(message.into())
    }
}

impl From<protobuf::tsp::MessageInfo> for MessageInfo {
    fn from(value: protobuf::tsp::MessageInfo) -> Self {
        Self {
            type_: value.r#type,
            versions: value.version,
            length: value.length,
            field_infos: value.field_infos,
            object_references: value.object_references,
            data_references: value.data_references,
            base_message_index: value.base_message_index,
            diff_merge_version: value.diff_merge_version,
            diff_field_path: value.diff_field_path,
            fields_to_remove: value.fields_to_remove,
            diff_read_version: value.diff_read_version,
        }
    }
}

impl From<&MessageInfo> for protobuf::tsp::MessageInfo {
    fn from(value: &MessageInfo) -> Self {
        Self {
            r#type: value.type_,
            version: value.versions.clone(),
            length: value.length,
            field_infos: value.field_infos.clone(),
            object_references: value.object_references.clone(),
            data_references: value.data_references.clone(),
            base_message_index: value.base_message_index,
            diff_merge_version: value.diff_merge_version.clone(),
            diff_field_path: value.diff_field_path.clone(),
            fields_to_remove: value.fields_to_remove.clone(),
            diff_read_version: value.diff_read_version.clone(),
        }
    }
}

/// A parsed, mutable IWA component archive.
#[derive(Debug)]
pub struct Archive {
    pub objects: Vec<ArchiveObject>,
}

impl Archive {
    pub fn new() -> Self {
        Self {
            objects: Vec::new(),
        }
    }

    /// Parse a decompressed IWA stream.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut objects = Vec::new();
        let mut cursor = Cursor::new(data);

        while cursor.position() < data.len() as u64 {
            let varint_start = cursor.position();
            let header_length = usize::try_from(varint::decode_varint(&mut cursor)?)
                .map_err(|_| Error::Archive("ArchiveInfo length exceeds usize".to_string()))?;
            let header_start = cursor.position();
            let varint_length = header_start - varint_start;
            let header_start_usize = usize::try_from(header_start)
                .map_err(|_| Error::Archive("ArchiveInfo offset exceeds usize".to_owned()))?;
            let header_end = header_start_usize
                .checked_add(header_length)
                .ok_or_else(|| Error::Archive("ArchiveInfo range overflow".to_owned()))?;
            let header_data = data.get(header_start_usize..header_end).ok_or_else(|| {
                Error::Archive(format!(
                    "ArchiveInfo declares {header_length} bytes but the IWA stream is truncated"
                ))
            })?;
            let archive_info: ArchiveInfo = protobuf::tsp::ArchiveInfo::decode(header_data)?.into();
            if archive_info.identifier.is_none() {
                return Err(Error::Archive(
                    "IWA object is missing its archive identifier".to_string(),
                ));
            }
            cursor.set_position(header_end as u64);
            let data_start = cursor.position();

            let payload_length =
                archive_info
                    .message_infos
                    .iter()
                    .try_fold(0usize, |total, info| {
                        total.checked_add(info.length as usize).ok_or_else(|| {
                            Error::Archive("IWA object payload length overflow".to_owned())
                        })
                    })?;
            let payload_start = usize::try_from(data_start)
                .map_err(|_| Error::Archive("IWA payload offset exceeds usize".to_owned()))?;
            let payload_end = payload_start
                .checked_add(payload_length)
                .ok_or_else(|| Error::Archive("IWA object payload range overflow".to_owned()))?;
            if payload_end > data.len() {
                return Err(Error::Archive(format!(
                    "IWA object declares {payload_length} payload bytes but the stream is truncated"
                )));
            }

            let mut messages = Vec::with_capacity(archive_info.message_infos.len());
            for info in &archive_info.message_infos {
                let mut message_data = vec![0; info.length as usize];
                cursor.read_exact(&mut message_data)?;
                messages.push(RawMessage {
                    type_: info.type_,
                    data: message_data,
                });
            }
            let data_length = cursor.position() - data_start;
            let decoded_messages = decode_raw_messages(&messages);

            objects.push(ArchiveObject {
                archive_info,
                messages,
                decoded_messages,
                header_offset: varint_start,
                header_length: varint_length + header_length as u64,
                data_offset: data_start,
                data_length,
            });
        }

        let archive = Self { objects };
        archive.validate()?;
        Ok(archive)
    }

    /// Serialize this archive into the decompressed IWA stream representation.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let estimated = self.objects.iter().try_fold(0usize, |total, object| {
            let payload = object
                .messages
                .iter()
                .try_fold(0usize, |sum, message| sum.checked_add(message.data.len()))?;
            total.checked_add(payload)?.checked_add(64)
        });
        let mut output = Vec::with_capacity(estimated.unwrap_or(0));

        for object in &self.objects {
            let mut info = object.archive_info.clone();
            for (message_info, message) in info.message_infos.iter_mut().zip(&object.messages) {
                message_info.type_ = message.type_;
                message_info.length = u32::try_from(message.data.len()).map_err(|_| {
                    Error::Archive("IWA message payload exceeds the u32 format limit".to_string())
                })?;
            }
            let header = info.encode();
            output.extend(varint::encode_varint(header.len() as u64));
            output.extend_from_slice(&header);
            for message in &object.messages {
                output.extend_from_slice(&message.data);
            }
        }
        Ok(output)
    }

    /// Validate object identifiers and message metadata before serialization.
    pub fn validate(&self) -> Result<()> {
        let mut identifiers = HashSet::with_capacity(self.objects.len());
        for object in &self.objects {
            let identifier = object.archive_info.identifier.ok_or_else(|| {
                Error::Archive("IWA object is missing its archive identifier".to_string())
            })?;
            if !identifiers.insert(identifier) {
                return Err(Error::Archive(format!(
                    "Duplicate IWA object identifier {identifier}"
                )));
            }
            if object.archive_info.message_infos.len() != object.messages.len() {
                return Err(Error::Archive(format!(
                    "IWA object {identifier} has {} MessageInfo entries for {} payloads",
                    object.archive_info.message_infos.len(),
                    object.messages.len()
                )));
            }
        }
        Ok(())
    }

    pub fn object(&self, identifier: u64) -> Option<&ArchiveObject> {
        self.objects
            .iter()
            .find(|object| object.archive_info.identifier == Some(identifier))
    }

    pub fn object_mut(&mut self, identifier: u64) -> Option<&mut ArchiveObject> {
        self.objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(identifier))
    }

    /// Insert a new object, rejecting duplicate identifiers.
    pub fn insert_object(&mut self, object: ArchiveObject) -> Result<()> {
        let identifier = object.archive_info.identifier.ok_or_else(|| {
            Error::Archive("Cannot insert an IWA object without an identifier".to_string())
        })?;
        if self.object(identifier).is_some() {
            return Err(Error::Archive(format!(
                "IWA object {identifier} already exists"
            )));
        }
        object.validate()?;
        self.objects.push(object);
        Ok(())
    }

    /// Insert or replace an object, returning the previous value when present.
    pub fn upsert_object(&mut self, object: ArchiveObject) -> Result<Option<ArchiveObject>> {
        let identifier = object.archive_info.identifier.ok_or_else(|| {
            Error::Archive("Cannot upsert an IWA object without an identifier".to_string())
        })?;
        object.validate()?;
        if let Some(index) = self
            .objects
            .iter()
            .position(|item| item.archive_info.identifier == Some(identifier))
        {
            Ok(Some(std::mem::replace(&mut self.objects[index], object)))
        } else {
            self.objects.push(object);
            Ok(None)
        }
    }

    /// Remove an object by identifier.
    pub fn remove_object(&mut self, identifier: u64) -> Option<ArchiveObject> {
        let index = self
            .objects
            .iter()
            .position(|object| object.archive_info.identifier == Some(identifier))?;
        Some(self.objects.remove(index))
    }
}

impl Default for Archive {
    fn default() -> Self {
        Self::new()
    }
}

/// One object within an IWA archive.
#[derive(Debug)]
pub struct ArchiveObject {
    pub archive_info: ArchiveInfo,
    pub messages: Vec<RawMessage>,
    pub decoded_messages: Vec<Box<dyn DecodedMessage>>,
    pub header_offset: u64,
    pub header_length: u64,
    pub data_offset: u64,
    pub data_length: u64,
}

impl ArchiveObject {
    /// Construct a new object from raw protobuf payloads.
    pub fn new(identifier: u64, messages: Vec<RawMessage>) -> Result<Self> {
        let message_infos = messages
            .iter()
            .map(|message| {
                let length = u32::try_from(message.data.len()).map_err(|_| {
                    Error::Archive("IWA message payload exceeds the u32 format limit".to_string())
                })?;
                Ok(MessageInfo::new(message.type_, length))
            })
            .collect::<Result<Vec<_>>>()?;
        let decoded_messages = decode_raw_messages(&messages);
        Ok(Self {
            archive_info: ArchiveInfo {
                identifier: Some(identifier),
                message_infos,
                should_merge: None,
            },
            messages,
            decoded_messages,
            header_offset: 0,
            header_length: 0,
            data_offset: 0,
            data_length: 0,
        })
    }

    pub fn validate(&self) -> Result<()> {
        if self.archive_info.identifier.is_none() {
            return Err(Error::Archive(
                "IWA object is missing its archive identifier".to_string(),
            ));
        }
        if self.archive_info.message_infos.len() != self.messages.len() {
            return Err(Error::Archive(
                "IWA object message metadata and payload counts differ".to_string(),
            ));
        }
        Ok(())
    }

    /// Replace a payload and keep its `MessageInfo` and decoded cache in sync.
    pub fn replace_message(&mut self, index: usize, message: RawMessage) -> Result<RawMessage> {
        let old = self
            .messages
            .get_mut(index)
            .ok_or_else(|| Error::Archive(format!("IWA message index {index} is out of bounds")))?;
        let length = u32::try_from(message.data.len()).map_err(|_| {
            Error::Archive("IWA message payload exceeds the u32 format limit".to_string())
        })?;
        let old = std::mem::replace(old, message);
        self.archive_info.message_infos[index].type_ = self.messages[index].type_;
        self.archive_info.message_infos[index].length = length;
        self.decoded_messages = decode_raw_messages(&self.messages);
        Ok(old)
    }

    pub fn push_message(&mut self, message: RawMessage) -> Result<()> {
        let length = u32::try_from(message.data.len()).map_err(|_| {
            Error::Archive("IWA message payload exceeds the u32 format limit".to_string())
        })?;
        self.archive_info
            .message_infos
            .push(MessageInfo::new(message.type_, length));
        self.messages.push(message);
        self.decoded_messages = decode_raw_messages(&self.messages);
        Ok(())
    }

    pub fn remove_message(&mut self, index: usize) -> Option<RawMessage> {
        if index >= self.messages.len() {
            return None;
        }
        self.archive_info.message_infos.remove(index);
        let message = self.messages.remove(index);
        self.decoded_messages = decode_raw_messages(&self.messages);
        Some(message)
    }

    pub fn extract_text(&self) -> Vec<String> {
        self.decoded_messages
            .iter()
            .flat_map(|message| message.extract_text())
            .collect()
    }

    pub fn primary_message_type(&self) -> Option<u32> {
        self.messages.first().map(|message| message.type_)
    }
}

/// Raw protobuf payload data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawMessage {
    pub type_: u32,
    pub data: Vec<u8>,
}

fn decode_raw_messages(messages: &[RawMessage]) -> Vec<Box<dyn DecodedMessage>> {
    let mut decoded_messages = Vec::new();
    for message in messages {
        if let Ok(decoded) = decode(message.type_, &message.data) {
            decoded_messages.push(decoded);
        } else if let Ok(storage) = protobuf::tswp::StorageArchive::decode(message.data.as_slice())
        {
            decoded_messages
                .push(Box::new(protobuf::StorageArchiveWrapper(storage)) as Box<dyn DecodedMessage>);
        }
    }
    decoded_messages
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn archive_round_trip_preserves_message_info() {
        let mut object = ArchiveObject::new(
            123,
            vec![RawMessage {
                type_: 42,
                data: vec![1, 2, 3],
            }],
        )
        .unwrap();
        object.archive_info.should_merge = Some(true);
        object.archive_info.message_infos[0].versions = vec![1, 2, 3];
        object.archive_info.message_infos[0].object_references = vec![99];
        object.archive_info.message_infos[0].data_references = vec![7];
        let archive = Archive {
            objects: vec![object],
        };

        let bytes = archive.to_bytes().unwrap();
        let reparsed = Archive::parse(&bytes).unwrap();
        let object = reparsed.object(123).unwrap();
        assert_eq!(object.messages[0].data, [1, 2, 3]);
        assert_eq!(object.archive_info.should_merge, Some(true));
        assert_eq!(object.archive_info.message_infos[0].versions, [1, 2, 3]);
        assert_eq!(object.archive_info.message_infos[0].object_references, [99]);
        assert_eq!(object.archive_info.message_infos[0].data_references, [7]);
    }

    #[test]
    fn archive_object_crud() {
        let mut archive = Archive::new();
        archive
            .insert_object(ArchiveObject::new(1, Vec::new()).unwrap())
            .unwrap();
        assert!(archive.object(1).is_some());
        assert!(
            archive
                .insert_object(ArchiveObject::new(1, Vec::new()).unwrap())
                .is_err()
        );

        let replacement = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 100,
                data: vec![8],
            }],
        )
        .unwrap();
        assert!(archive.upsert_object(replacement).unwrap().is_some());
        assert_eq!(archive.object(1).unwrap().primary_message_type(), Some(100));
        assert!(archive.remove_object(1).is_some());
        assert!(archive.object(1).is_none());
    }

    #[test]
    fn rejects_declared_lengths_before_allocating_payloads() {
        assert!(Archive::parse(&[127]).is_err());

        let header = protobuf::tsp::ArchiveInfo {
            identifier: Some(1),
            message_infos: vec![protobuf::tsp::MessageInfo {
                r#type: 99,
                length: 1_000_000_000,
                ..Default::default()
            }],
            should_merge: None,
        }
        .encode_to_vec();
        let mut bytes = varint::encode_varint(header.len() as u64);
        bytes.extend_from_slice(&header);
        assert!(Archive::parse(&bytes).is_err());
    }

    #[test]
    fn rejects_missing_identifier_before_materializing_later_objects() {
        // An empty ArchiveInfo has no identifier. The trailing zero-length
        // headers would otherwise be materialized as objects before the
        // malformed varint is reached.
        let mut bytes = vec![0];
        bytes.extend(std::iter::repeat(0).take(4096));
        bytes.push(0x80);

        let error = Archive::parse(&bytes).unwrap_err();

        assert!(matches!(
            error,
            Error::Archive(message)
                if message == "IWA object is missing its archive identifier"
        ));
    }
}
