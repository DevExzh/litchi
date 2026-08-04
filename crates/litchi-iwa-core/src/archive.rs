//! Bounded framing for decompressed IWA archive streams.
//!
//! An IWA stream is a sequence of objects. Each object contains a varint
//! length, a `TSP.ArchiveInfo` protobuf, and the payloads described by that
//! header's `TSP.MessageInfo` entries. This module owns only that physical
//! framing. It does not resolve message types or decode application payloads.

#![allow(
    clippy::cast_possible_truncation,
    reason = "Varint encoding masks each narrowed byte to its seven-bit wire value."
)]
#![allow(
    clippy::map_err_ignore,
    reason = "The public physical-layer error API intentionally normalizes conversion and allocation failures."
)]
#![allow(
    clippy::missing_errors_doc,
    reason = "The module's Result-returning methods share the crate-level physical framing error contract."
)]
#![allow(
    clippy::module_name_repetitions,
    reason = "These names mirror the neutral IWA wire vocabulary and remain source-compatible with the extracted layer."
)]
#![allow(
    clippy::shadow_reuse,
    reason = "Validated limit profiles deliberately retain the public parameter name."
)]

use litchi_iwa_protos::tsp;
use prost::Message;

use crate::{Error, LimitKind, Limits, Result};

const MAX_VARINT_BYTES: usize = 10;

/// Metadata for one object in an IWA component.
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
    /// Decode one bounded `TSP.ArchiveInfo` header.
    pub fn decode(data: &[u8]) -> Result<Self> {
        Self::decode_with_limits(data, Limits::default())
    }

    /// Decode one `TSP.ArchiveInfo` header under explicit limits.
    pub fn decode_with_limits(data: &[u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        check_header_length(data.len(), limits)?;
        let archive_info = Self::from(tsp::ArchiveInfo::decode(data)?);
        archive_info.validate_with_limits(limits)?;
        Ok(archive_info)
    }

    /// Construct metadata for an object from its raw payload metadata.
    #[must_use]
    pub fn new(identifier: u64, message_infos: Vec<MessageInfo>) -> Self {
        Self {
            identifier: Some(identifier),
            message_infos,
            should_merge: None,
        }
    }

    fn validate_with_limits(&self, limits: Limits) -> Result<()> {
        let message_count = self.message_infos.len();
        if message_count > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                message_count,
                limits.max_messages_per_object(),
            ));
        }

        let mut metadata_items = message_count;
        for message_info in &self.message_infos {
            message_info.validate_with_limits(limits)?;
            add_limited(
                &mut metadata_items,
                message_info.metadata_item_count()?,
                LimitKind::MetadataItems,
                limits.max_metadata_items(),
            )?;
        }
        if self.identifier.is_none() {
            return Err(Error::invalid_archive(
                0,
                "object is missing its archive identifier",
            ));
        }
        Ok(())
    }
}

impl From<tsp::ArchiveInfo> for ArchiveInfo {
    fn from(value: tsp::ArchiveInfo) -> Self {
        Self {
            identifier: value.identifier,
            message_infos: value.message_infos.into_iter().map(Into::into).collect(),
            should_merge: value.should_merge,
        }
    }
}

impl From<&ArchiveInfo> for tsp::ArchiveInfo {
    fn from(value: &ArchiveInfo) -> Self {
        Self {
            identifier: value.identifier,
            message_infos: value.message_infos.iter().map(Into::into).collect(),
            should_merge: value.should_merge,
        }
    }
}

/// Metadata for one payload in an archive object.
#[derive(Debug, Clone, PartialEq)]
pub struct MessageInfo {
    pub type_: u32,
    pub versions: Vec<u32>,
    pub length: u32,
    pub field_infos: Vec<tsp::FieldInfo>,
    pub object_references: Vec<u64>,
    pub data_references: Vec<u64>,
    pub base_message_index: Option<u32>,
    pub diff_merge_version: Vec<u32>,
    pub diff_field_path: Option<tsp::FieldPath>,
    pub fields_to_remove: Vec<tsp::FieldPath>,
    pub diff_read_version: Vec<u32>,
}

impl MessageInfo {
    /// Construct the conventional metadata for a newly created payload.
    #[must_use]
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

    /// Decode one bounded `TSP.MessageInfo` protobuf.
    pub fn decode(data: &[u8]) -> Result<Self> {
        Self::decode_with_limits(data, Limits::default())
    }

    /// Decode one `TSP.MessageInfo` protobuf under explicit limits.
    pub fn decode_with_limits(data: &[u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        check_header_length(data.len(), limits)?;
        let message_info = Self::from(tsp::MessageInfo::decode(data)?);
        message_info.validate_with_limits(limits)?;
        Ok(message_info)
    }

    fn validate_with_limits(&self, limits: Limits) -> Result<()> {
        let length = usize::try_from(self.length)
            .map_err(|_| Error::invalid_archive(0, "message length exceeds usize"))?;
        if length > limits.max_message_bytes() {
            return Err(limit(
                LimitKind::MessageBytes,
                length,
                limits.max_message_bytes(),
            ));
        }
        let metadata_items = self.metadata_item_count()?;
        if metadata_items > limits.max_metadata_items() {
            return Err(limit(
                LimitKind::MetadataItems,
                metadata_items,
                limits.max_metadata_items(),
            ));
        }
        Ok(())
    }

    fn metadata_item_count(&self) -> Result<usize> {
        let mut count = 0;
        add_count(&mut count, self.versions.len())?;
        add_count(&mut count, self.field_infos.len())?;
        add_count(&mut count, self.object_references.len())?;
        add_count(&mut count, self.data_references.len())?;
        add_count(&mut count, self.diff_merge_version.len())?;
        add_count(&mut count, usize::from(self.diff_field_path.is_some()))?;
        add_count(&mut count, self.fields_to_remove.len())?;
        add_count(&mut count, self.diff_read_version.len())?;
        for field_info in &self.field_infos {
            add_count(&mut count, 1)?;
            add_count(&mut count, field_info.path.path.len())?;
            add_count(&mut count, field_info.object_references.len())?;
            add_count(&mut count, field_info.data_references.len())?;
            add_count(&mut count, field_info.known_field_version.len())?;
            add_count(
                &mut count,
                usize::from(field_info.known_field_feature_identifier.is_some()),
            )?;
        }
        if let Some(path) = &self.diff_field_path {
            add_count(&mut count, path.path.len())?;
        }
        for path in &self.fields_to_remove {
            add_count(&mut count, path.path.len())?;
        }
        Ok(count)
    }
}

impl From<tsp::MessageInfo> for MessageInfo {
    fn from(value: tsp::MessageInfo) -> Self {
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

impl From<&MessageInfo> for tsp::MessageInfo {
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

/// Raw payload data for one archive message.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawMessage {
    pub type_: u32,
    pub data: Vec<u8>,
}

/// One object within a decompressed IWA archive stream.
#[derive(Debug, Clone, PartialEq)]
pub struct ArchiveObject {
    pub archive_info: ArchiveInfo,
    pub messages: Vec<RawMessage>,
    /// Offset of the object's length prefix in the decompressed stream.
    pub header_offset: u64,
    /// Size of the length prefix and encoded `ArchiveInfo` header.
    pub header_length: u64,
    /// Offset of the first payload in the decompressed stream.
    pub data_offset: u64,
    /// Total bytes occupied by the object's payloads.
    pub data_length: u64,
}

impl ArchiveObject {
    /// Construct an object and conventional metadata for its raw payloads.
    pub fn new(identifier: u64, messages: Vec<RawMessage>) -> Result<Self> {
        Self::new_with_limits(identifier, messages, Limits::default())
    }

    /// Construct an object under explicit resource limits.
    pub fn new_with_limits(
        identifier: u64,
        messages: Vec<RawMessage>,
        limits: Limits,
    ) -> Result<Self> {
        let limits = limits.validate()?;
        if messages.len() > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                messages.len(),
                limits.max_messages_per_object(),
            ));
        }
        let mut message_infos = Vec::new();
        message_infos
            .try_reserve_exact(messages.len())
            .map_err(|_| Error::allocation("IWA message metadata", messages.len()))?;
        for message in &messages {
            let length = u32::try_from(message.data.len())
                .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
            check_message_length(message.data.len(), limits)?;
            message_infos.push(MessageInfo::new(message.type_, length));
        }
        let object = Self {
            archive_info: ArchiveInfo::new(identifier, message_infos),
            messages,
            header_offset: 0,
            header_length: 0,
            data_offset: 0,
            data_length: 0,
        };
        object.validate_with_limits(limits)?;
        Ok(object)
    }

    /// Validate object metadata, payload sizes, and encoded header size.
    pub fn validate(&self) -> Result<()> {
        self.validate_with_limits(Limits::default())
    }

    /// Validate this object under explicit resource limits.
    pub fn validate_with_limits(&self, limits: Limits) -> Result<()> {
        let limits = limits.validate()?;
        if self.archive_info.identifier.is_none() {
            return Err(Error::invalid_archive(
                0,
                "object is missing its archive identifier",
            ));
        }
        if self.archive_info.message_infos.len() != self.messages.len() {
            return Err(Error::invalid_archive(
                0,
                "message metadata and payload counts differ",
            ));
        }
        self.archive_info.validate_with_limits(limits)?;
        let header_length = tsp::ArchiveInfo::from(&self.archive_info).encoded_len();
        check_header_length(header_length, limits)?;
        let prefix_length = varint_len(header_length)?;
        let mut payload_length = 0;
        for message in &self.messages {
            check_message_length(message.data.len(), limits)?;
            add_limited(
                &mut payload_length,
                message.data.len(),
                LimitKind::ObjectBytes,
                limits.max_object_bytes(),
            )?;
        }
        let object_length = prefix_length
            .checked_add(header_length)
            .and_then(|length| length.checked_add(payload_length))
            .ok_or_else(|| Error::invalid_archive(0, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
        Ok(())
    }

    /// Return the type identifier of the first payload, if present.
    #[must_use]
    pub fn primary_message_type(&self) -> Option<u32> {
        self.messages.first().map(|message| message.type_)
    }
}

/// A parsed, mutable decompressed IWA component archive.
#[derive(Debug, Clone, PartialEq)]
pub struct Archive {
    pub objects: Vec<ArchiveObject>,
}

impl Archive {
    /// Construct an empty archive.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            objects: Vec::new(),
        }
    }

    /// Parse a decompressed IWA stream.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Self::parse_with_limits(data, Limits::default())
    }

    /// Parse a decompressed IWA stream under explicit resource limits.
    pub fn parse_with_limits(data: &[u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        if data.len() > limits.max_archive_bytes() {
            return Err(limit(
                LimitKind::ArchiveBytes,
                data.len(),
                limits.max_archive_bytes(),
            ));
        }

        let mut objects = Vec::new();
        objects
            .try_reserve(data.len().min(limits.max_objects()))
            .map_err(|_| Error::allocation("IWA archive objects", data.len()))?;
        let mut cursor = 0usize;
        let mut total_messages = 0usize;
        while cursor < data.len() {
            let object_start = cursor;
            if objects.len() >= limits.max_objects() {
                return Err(limit(
                    LimitKind::Objects,
                    objects.len() + 1,
                    limits.max_objects(),
                ));
            }
            let (header_length_u64, prefix_length) = decode_varint(&data[cursor..])?;
            cursor = cursor
                .checked_add(prefix_length)
                .ok_or_else(|| Error::invalid_archive(object_start, "header offset overflow"))?;
            let header_length = usize::try_from(header_length_u64)
                .map_err(|_| Error::invalid_archive(object_start, "header length exceeds usize"))?;
            check_header_length(header_length, limits)?;
            let header_end = cursor
                .checked_add(header_length)
                .ok_or_else(|| Error::invalid_archive(object_start, "header range overflow"))?;
            let header = data.get(cursor..header_end).ok_or_else(|| {
                Error::invalid_archive(object_start, "truncated ArchiveInfo header")
            })?;
            let archive_info = ArchiveInfo::decode_with_limits(header, limits)?;
            cursor = header_end;
            add_limited(
                &mut total_messages,
                archive_info.message_infos.len(),
                LimitKind::Messages,
                limits.max_messages(),
            )?;

            let object_prefix = prefix_length
                .checked_add(header_length)
                .ok_or_else(|| Error::invalid_archive(object_start, "object prefix overflow"))?;
            if object_prefix > limits.max_object_bytes() {
                return Err(limit(
                    LimitKind::ObjectBytes,
                    object_prefix,
                    limits.max_object_bytes(),
                ));
            }
            let mut payload_length = 0usize;
            for message_info in &archive_info.message_infos {
                let length = usize::try_from(message_info.length)
                    .map_err(|_| Error::invalid_archive(cursor, "message length exceeds usize"))?;
                add_limited(
                    &mut payload_length,
                    length,
                    LimitKind::ObjectBytes,
                    limits.max_object_bytes() - object_prefix,
                )?;
            }
            let payload_end = cursor
                .checked_add(payload_length)
                .ok_or_else(|| Error::invalid_archive(cursor, "payload range overflow"))?;
            if payload_end > data.len() {
                return Err(Error::invalid_archive(cursor, "truncated message payload"));
            }

            let mut messages = Vec::new();
            messages
                .try_reserve_exact(archive_info.message_infos.len())
                .map_err(|_| {
                    Error::allocation("IWA object messages", archive_info.message_infos.len())
                })?;
            let payload_start = cursor;
            for message_info in &archive_info.message_infos {
                let length = usize::try_from(message_info.length)
                    .map_err(|_| Error::invalid_archive(cursor, "message length exceeds usize"))?;
                let end = cursor
                    .checked_add(length)
                    .ok_or_else(|| Error::invalid_archive(cursor, "message range overflow"))?;
                let message_data = data
                    .get(cursor..end)
                    .ok_or_else(|| Error::invalid_archive(cursor, "truncated message payload"))?;
                let mut owned = Vec::new();
                owned
                    .try_reserve_exact(length)
                    .map_err(|_| Error::allocation("IWA message payload", length))?;
                owned.extend_from_slice(message_data);
                messages.push(RawMessage {
                    type_: message_info.type_,
                    data: owned,
                });
                cursor = end;
            }
            objects.push(ArchiveObject {
                archive_info,
                messages,
                header_offset: u64::try_from(object_start)
                    .map_err(|_| Error::invalid_archive(object_start, "offset exceeds u64"))?,
                header_length: u64::try_from(prefix_length + header_length)
                    .map_err(|_| Error::invalid_archive(object_start, "header exceeds u64"))?,
                data_offset: u64::try_from(payload_start)
                    .map_err(|_| Error::invalid_archive(payload_start, "offset exceeds u64"))?,
                data_length: u64::try_from(payload_length)
                    .map_err(|_| Error::invalid_archive(payload_start, "payload exceeds u64"))?,
            });
        }
        Ok(Self { objects })
    }

    /// Serialize this archive as a decompressed IWA stream.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_bytes_with_limits(Limits::default())
    }

    /// Serialize this archive under explicit resource limits.
    pub fn to_bytes_with_limits(&self, limits: Limits) -> Result<Vec<u8>> {
        let limits = limits.validate()?;
        if self.objects.len() > limits.max_objects() {
            return Err(limit(
                LimitKind::Objects,
                self.objects.len(),
                limits.max_objects(),
            ));
        }
        let mut output = Vec::new();
        for object in &self.objects {
            object.validate_with_limits(limits)?;
            let mut info = object.archive_info.clone();
            for (message_info, message) in info.message_infos.iter_mut().zip(&object.messages) {
                message_info.type_ = message.type_;
                message_info.length = u32::try_from(message.data.len())
                    .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
            }
            let encoded_info = tsp::ArchiveInfo::from(&info);
            let header_length = encoded_info.encoded_len();
            check_header_length(header_length, limits)?;
            let mut header = Vec::new();
            header
                .try_reserve_exact(header_length)
                .map_err(|_| Error::allocation("IWA ArchiveInfo header", header_length))?;
            encoded_info.encode(&mut header)?;
            let prefix_length = varint_len(header_length)?;
            let payload_length = object.messages.iter().try_fold(0usize, |total, message| {
                total
                    .checked_add(message.data.len())
                    .ok_or_else(|| Error::invalid_archive(0, "payload length overflow"))
            })?;
            let object_length = prefix_length
                .checked_add(header_length)
                .and_then(|length| length.checked_add(payload_length))
                .ok_or_else(|| Error::invalid_archive(0, "object length overflow"))?;
            if object_length > limits.max_object_bytes() {
                return Err(limit(
                    LimitKind::ObjectBytes,
                    object_length,
                    limits.max_object_bytes(),
                ));
            }
            let next_length = output
                .len()
                .checked_add(object_length)
                .ok_or_else(|| Error::invalid_archive(0, "archive length overflow"))?;
            if next_length > limits.max_archive_bytes() {
                return Err(limit(
                    LimitKind::ArchiveBytes,
                    next_length,
                    limits.max_archive_bytes(),
                ));
            }
            output
                .try_reserve_exact(object_length)
                .map_err(|_| Error::allocation("IWA archive output", object_length))?;
            let mut prefix = [0u8; MAX_VARINT_BYTES];
            let prefix = encode_varint(
                u64::try_from(header_length)
                    .map_err(|_| Error::invalid_archive(0, "header length exceeds u64"))?,
                &mut prefix,
            );
            output.extend_from_slice(prefix);
            output.extend_from_slice(&header);
            for message in &object.messages {
                output.extend_from_slice(&message.data);
            }
        }
        Ok(output)
    }

    /// Validate all objects under the default resource limits.
    pub fn validate(&self) -> Result<()> {
        self.validate_with_limits(Limits::default())
    }

    /// Validate all objects under explicit resource limits.
    pub fn validate_with_limits(&self, limits: Limits) -> Result<()> {
        let limits = limits.validate()?;
        if self.objects.len() > limits.max_objects() {
            return Err(limit(
                LimitKind::Objects,
                self.objects.len(),
                limits.max_objects(),
            ));
        }
        let mut total_messages = 0usize;
        let mut total_payload = 0usize;
        for object in &self.objects {
            object.validate_with_limits(limits)?;
            add_limited(
                &mut total_messages,
                object.messages.len(),
                LimitKind::Messages,
                limits.max_messages(),
            )?;
            for message in &object.messages {
                add_limited(
                    &mut total_payload,
                    message.data.len(),
                    LimitKind::ArchiveBytes,
                    limits.max_archive_bytes(),
                )?;
            }
        }
        Ok(())
    }

    /// Insert an object, rejecting missing or duplicate identifiers.
    pub fn insert_object(&mut self, object: ArchiveObject) -> Result<()> {
        self.insert_object_with_limits(object, Limits::default())
    }

    /// Insert an object under explicit resource limits.
    pub fn insert_object_with_limits(
        &mut self,
        object: ArchiveObject,
        limits: Limits,
    ) -> Result<()> {
        let limits = limits.validate()?;
        let identifier = object
            .archive_info
            .identifier
            .ok_or_else(|| Error::invalid_archive(0, "object is missing its archive identifier"))?;
        if self
            .objects
            .iter()
            .any(|current| current.archive_info.identifier == Some(identifier))
        {
            return Err(Error::invalid_archive(0, "duplicate object identifier"));
        }
        self.objects
            .try_reserve(1)
            .map_err(|_| Error::allocation("IWA archive objects", 1))?;
        self.objects.push(object);
        if let Err(error) = self.validate_with_limits(limits) {
            self.objects.pop();
            return Err(error);
        }
        Ok(())
    }
}

impl Default for Archive {
    fn default() -> Self {
        Self::new()
    }
}

fn check_header_length(length: usize, limits: Limits) -> Result<()> {
    if length > limits.max_header_bytes() {
        return Err(limit(
            LimitKind::HeaderBytes,
            length,
            limits.max_header_bytes(),
        ));
    }
    Ok(())
}

fn check_message_length(length: usize, limits: Limits) -> Result<()> {
    if length > limits.max_message_bytes() {
        return Err(limit(
            LimitKind::MessageBytes,
            length,
            limits.max_message_bytes(),
        ));
    }
    Ok(())
}

fn add_count(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| Error::invalid_archive(0, "metadata count overflow"))?;
    Ok(())
}

fn add_limited(total: &mut usize, amount: usize, kind: LimitKind, maximum: usize) -> Result<()> {
    let observed = total
        .checked_add(amount)
        .ok_or_else(|| Error::invalid_archive(0, "resource count overflow"))?;
    if observed > maximum {
        return Err(limit(kind, observed, maximum));
    }
    *total = observed;
    Ok(())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}

fn decode_varint(data: &[u8]) -> Result<(u64, usize)> {
    let mut value = 0u64;
    for (index, byte) in data.iter().copied().take(MAX_VARINT_BYTES).enumerate() {
        if index == MAX_VARINT_BYTES - 1 && byte > 1 {
            return Err(Error::invalid_archive(
                index,
                "archive length varint overflow",
            ));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Ok((value, index + 1));
        }
    }
    if data.len() < MAX_VARINT_BYTES {
        Err(Error::invalid_archive(
            data.len(),
            "truncated archive length varint",
        ))
    } else {
        Err(Error::invalid_archive(
            MAX_VARINT_BYTES - 1,
            "archive length varint overflow",
        ))
    }
}

fn varint_len(value: usize) -> Result<usize> {
    let value =
        u64::try_from(value).map_err(|_| Error::invalid_archive(0, "length exceeds u64"))?;
    Ok(if value == 0 {
        1
    } else {
        (u64::BITS as usize - value.leading_zeros() as usize).div_ceil(7)
    })
}

fn encode_varint(mut value: u64, output: &mut [u8; MAX_VARINT_BYTES]) -> &[u8] {
    let mut length = 0;
    while value >= 0x80 {
        output[length] = u8::try_from(value & 0x7f).unwrap_or_default() | 0x80;
        value >>= 7;
        length += 1;
    }
    output[length] = u8::try_from(value).unwrap_or_default();
    &output[..=length]
}
