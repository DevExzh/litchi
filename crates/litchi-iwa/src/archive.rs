//! Metadata-preserving IWA archive parsing, mutation, and serialization.
//!
//! A decompressed IWA stream is a sequence of objects. Each object starts with
//! a varint-sized `TSP.ArchiveInfo` protobuf followed by the payloads described
//! by its `TSP.MessageInfo` entries.

use std::collections::HashSet;
use std::io::{Cursor, Read};

use prost::Message;

use crate::protobuf::{self, DecodedMessage, decode};
use crate::snappy::SnappyStream;
use crate::varint;
use crate::{Error, Result};

/// Resource ceilings for one decompressed IWA archive.
///
/// The defaults are deliberately finite and fit below the Snappy stream
/// ceiling. Callers may tighten or raise an individual budget through the
/// builder methods, but the format-wide ceilings and cross-field invariants
/// are always enforced.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ArchiveLimits {
    max_archive_bytes: usize,
    max_objects: usize,
    max_messages: usize,
    max_messages_per_object: usize,
    max_object_bytes: usize,
    max_message_bytes: usize,
    max_header_bytes: usize,
    max_metadata_items: usize,
}

impl ArchiveLimits {
    /// Hard ceiling for one decompressed archive stream.
    pub const MAX_ARCHIVE_BYTES: usize = SnappyStream::MAX_DECOMPRESSED_STREAM;
    /// Hard ceiling for objects in one archive.
    pub const MAX_OBJECTS: usize = 100_000;
    /// Hard ceiling for messages across one archive.
    pub const MAX_MESSAGES: usize = 1_000_000;
    /// Hard ceiling for messages in one object.
    pub const MAX_MESSAGES_PER_OBJECT: usize = 100_000;
    /// Hard ceiling for one object, including its header and payloads.
    pub const MAX_OBJECT_BYTES: usize = SnappyStream::MAX_DECOMPRESSED_STREAM;
    /// Hard ceiling for one message payload.
    pub const MAX_MESSAGE_BYTES: usize = SnappyStream::MAX_DECOMPRESSED_STREAM;
    /// Hard ceiling for one encoded `ArchiveInfo` header.
    pub const MAX_HEADER_BYTES: usize = 16 * 1024 * 1024;
    /// Hard ceiling for repeated metadata items in one object header.
    pub const MAX_METADATA_ITEMS: usize = 1_000_000;

    /// Tighten or raise the aggregate archive byte budget.
    pub fn with_archive_bytes(mut self, value: usize) -> Result<Self> {
        check_configured_limit("archive byte", value, Self::MAX_ARCHIVE_BYTES)?;
        self.max_archive_bytes = value;
        self.max_object_bytes = self.max_object_bytes.min(value);
        self.max_message_bytes = self.max_message_bytes.min(value);
        self.max_header_bytes = self.max_header_bytes.min(value);
        self.validate()
    }

    /// Tighten or raise the object-count budget.
    pub fn with_objects(mut self, value: usize) -> Result<Self> {
        check_configured_limit("object", value, Self::MAX_OBJECTS)?;
        self.max_objects = value;
        self.validate()
    }

    /// Tighten or raise the aggregate message-count budget.
    pub fn with_messages(mut self, value: usize) -> Result<Self> {
        check_configured_limit("message", value, Self::MAX_MESSAGES)?;
        self.max_messages = value;
        self.max_messages_per_object = self.max_messages_per_object.min(value);
        self.validate()
    }

    /// Tighten or raise the per-object message-count budget.
    pub fn with_messages_per_object(mut self, value: usize) -> Result<Self> {
        check_configured_limit("per-object message", value, Self::MAX_MESSAGES_PER_OBJECT)?;
        self.max_messages_per_object = value;
        self.validate()
    }

    /// Tighten or raise the per-object wire-size budget.
    pub fn with_object_bytes(mut self, value: usize) -> Result<Self> {
        check_configured_limit("object byte", value, Self::MAX_OBJECT_BYTES)?;
        self.max_object_bytes = value;
        self.max_message_bytes = self.max_message_bytes.min(value);
        self.max_header_bytes = self.max_header_bytes.min(value);
        self.validate()
    }

    /// Tighten or raise the per-message payload budget.
    pub fn with_message_bytes(mut self, value: usize) -> Result<Self> {
        check_configured_limit("message byte", value, Self::MAX_MESSAGE_BYTES)?;
        self.max_message_bytes = value;
        self.validate()
    }

    /// Tighten or raise the per-header byte budget.
    pub fn with_header_bytes(mut self, value: usize) -> Result<Self> {
        check_configured_limit("header byte", value, Self::MAX_HEADER_BYTES)?;
        self.max_header_bytes = value;
        self.validate()
    }

    /// Tighten or raise the repeated-header-metadata budget.
    pub fn with_metadata_items(mut self, value: usize) -> Result<Self> {
        check_configured_limit("metadata item", value, Self::MAX_METADATA_ITEMS)?;
        self.max_metadata_items = value;
        self.validate()
    }

    /// Maximum decompressed bytes accepted for one archive.
    pub const fn max_archive_bytes(self) -> usize {
        self.max_archive_bytes
    }

    /// Maximum number of objects accepted for one archive.
    pub const fn max_objects(self) -> usize {
        self.max_objects
    }

    /// Maximum number of messages accepted for one archive.
    pub const fn max_messages(self) -> usize {
        self.max_messages
    }

    /// Maximum number of messages accepted for one object.
    pub const fn max_messages_per_object(self) -> usize {
        self.max_messages_per_object
    }

    /// Maximum wire bytes accepted for one object.
    pub const fn max_object_bytes(self) -> usize {
        self.max_object_bytes
    }

    /// Maximum payload bytes accepted for one message.
    pub const fn max_message_bytes(self) -> usize {
        self.max_message_bytes
    }

    /// Maximum encoded bytes accepted for one object header.
    pub const fn max_header_bytes(self) -> usize {
        self.max_header_bytes
    }

    /// Maximum repeated metadata items accepted in one object header.
    pub const fn max_metadata_items(self) -> usize {
        self.max_metadata_items
    }

    fn validate(self) -> Result<Self> {
        if self.max_archive_bytes == 0
            || self.max_objects == 0
            || self.max_messages == 0
            || self.max_messages_per_object == 0
            || self.max_object_bytes == 0
            || self.max_message_bytes == 0
            || self.max_header_bytes == 0
            || self.max_metadata_items == 0
        {
            return Err(Error::Archive(
                "IWA archive limits must be non-zero".to_owned(),
            ));
        }
        if self.max_messages_per_object > self.max_messages {
            return Err(Error::Archive(
                "IWA per-object message limit exceeds the aggregate message limit".to_owned(),
            ));
        }
        if self.max_object_bytes > self.max_archive_bytes {
            return Err(Error::Archive(
                "IWA object byte limit exceeds the aggregate archive byte limit".to_owned(),
            ));
        }
        if self.max_message_bytes > self.max_object_bytes {
            return Err(Error::Archive(
                "IWA message byte limit exceeds the object byte limit".to_owned(),
            ));
        }
        if self.max_header_bytes > self.max_object_bytes {
            return Err(Error::Archive(
                "IWA header byte limit exceeds the object byte limit".to_owned(),
            ));
        }
        Ok(self)
    }
}

impl Default for ArchiveLimits {
    fn default() -> Self {
        Self {
            max_archive_bytes: Self::MAX_ARCHIVE_BYTES,
            max_objects: Self::MAX_OBJECTS,
            max_messages: Self::MAX_MESSAGES,
            max_messages_per_object: Self::MAX_MESSAGES_PER_OBJECT,
            max_object_bytes: Self::MAX_OBJECT_BYTES,
            max_message_bytes: Self::MAX_MESSAGE_BYTES,
            max_header_bytes: Self::MAX_HEADER_BYTES,
            max_metadata_items: Self::MAX_METADATA_ITEMS,
        }
    }
}

fn check_configured_limit(resource: &str, value: usize, hard_limit: usize) -> Result<()> {
    if value == 0 {
        return Err(Error::Archive(format!(
            "IWA {resource} limit must be non-zero"
        )));
    }
    if value > hard_limit {
        return Err(Error::Archive(format!(
            "IWA {resource} limit {value} exceeds hard ceiling {hard_limit}"
        )));
    }
    Ok(())
}

fn limit_error(resource: &str, observed: usize, limit: usize, path: &str) -> Error {
    Error::Archive(format!(
        "IWA {resource} limit exceeded at {path}: observed {observed}, limit {limit}"
    ))
}

fn overflow_error(resource: &str, path: &str) -> Error {
    Error::Archive(format!("IWA {resource} arithmetic overflow at {path}"))
}

fn checked_add_limited(
    total: &mut usize,
    amount: usize,
    resource: &str,
    limit: usize,
    path: &str,
) -> Result<()> {
    let observed = total
        .checked_add(amount)
        .ok_or_else(|| overflow_error(resource, path))?;
    if observed > limit {
        return Err(limit_error(resource, observed, limit, path));
    }
    *total = observed;
    Ok(())
}

fn read_bounded<R: Read>(reader: &mut R, limit: usize, path: &str) -> Result<Vec<u8>> {
    let read_limit = u64::try_from(limit)
        .ok()
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| overflow_error("header byte", path))?;
    let mut data = Vec::new();
    reader.take(read_limit).read_to_end(&mut data)?;
    if data.len() > limit {
        return Err(limit_error("header byte", data.len(), limit, path));
    }
    Ok(data)
}

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
        Self::parse_with_limits(reader, ArchiveLimits::default())
    }

    /// Decode an `ArchiveInfo` protobuf under an explicit resource budget.
    pub fn parse_with_limits<R: Read>(reader: &mut R, limits: ArchiveLimits) -> Result<Self> {
        let limits = limits.validate()?;
        let data = read_bounded(reader, limits.max_header_bytes, "ArchiveInfo")?;
        let message = protobuf::tsp::ArchiveInfo::decode(data.as_slice())?;
        let archive_info: Self = message.into();
        archive_info.validate_with_limits(limits, "ArchiveInfo")?;
        Ok(archive_info)
    }

    fn validate_with_limits(&self, limits: ArchiveLimits, path: &str) -> Result<()> {
        let message_count = self.message_infos.len();
        if message_count > limits.max_messages_per_object {
            return Err(limit_error(
                "per-object message",
                message_count,
                limits.max_messages_per_object,
                path,
            ));
        }

        let mut metadata_items = message_count;
        for message_info in &self.message_infos {
            let message_length = usize::try_from(message_info.length).map_err(|_| {
                Error::Archive("IWA message length exceeds the platform usize limit".to_owned())
            })?;
            if message_length > limits.max_message_bytes {
                return Err(limit_error(
                    "message byte",
                    message_length,
                    limits.max_message_bytes,
                    path,
                ));
            }
            let count = message_info.metadata_item_count()?;
            checked_add_limited(
                &mut metadata_items,
                count,
                "metadata item",
                limits.max_metadata_items,
                path,
            )?;
        }
        Ok(())
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
        Self::parse_with_limits(reader, ArchiveLimits::default())
    }

    /// Decode a `MessageInfo` protobuf under an explicit resource budget.
    pub fn parse_with_limits<R: Read>(reader: &mut R, limits: ArchiveLimits) -> Result<Self> {
        let limits = limits.validate()?;
        let data = read_bounded(reader, limits.max_header_bytes, "MessageInfo")?;
        let message = protobuf::tsp::MessageInfo::decode(data.as_slice())?;
        let message_info: Self = message.into();
        let metadata_items = message_info.metadata_item_count()?;
        if metadata_items > limits.max_metadata_items {
            return Err(limit_error(
                "metadata item",
                metadata_items,
                limits.max_metadata_items,
                "MessageInfo",
            ));
        }
        let message_length = usize::try_from(message_info.length).map_err(|_| {
            Error::Archive("IWA message length exceeds the platform usize limit".to_owned())
        })?;
        if message_length > limits.max_message_bytes {
            return Err(limit_error(
                "message byte",
                message_length,
                limits.max_message_bytes,
                "MessageInfo.length",
            ));
        }
        Ok(message_info)
    }

    fn metadata_item_count(&self) -> Result<usize> {
        let mut count = 0;
        add_metadata_count(&mut count, self.versions.len(), "MessageInfo.version")?;
        add_metadata_count(
            &mut count,
            self.field_infos.len(),
            "MessageInfo.field_infos",
        )?;
        add_metadata_count(
            &mut count,
            self.object_references.len(),
            "MessageInfo.object_references",
        )?;
        add_metadata_count(
            &mut count,
            self.data_references.len(),
            "MessageInfo.data_references",
        )?;
        add_metadata_count(
            &mut count,
            self.diff_merge_version.len(),
            "MessageInfo.diff_merge_version",
        )?;
        add_metadata_count(
            &mut count,
            usize::from(self.diff_field_path.is_some()),
            "MessageInfo.diff_field_path",
        )?;
        add_metadata_count(
            &mut count,
            self.fields_to_remove.len(),
            "MessageInfo.fields_to_remove",
        )?;
        add_metadata_count(
            &mut count,
            self.diff_read_version.len(),
            "MessageInfo.diff_read_version",
        )?;
        for field_info in &self.field_infos {
            add_metadata_count(&mut count, 1, "FieldInfo.path")?;
            add_metadata_count(&mut count, field_info.path.path.len(), "FieldPath.path")?;
            add_metadata_count(
                &mut count,
                field_info.object_references.len(),
                "FieldInfo.object_references",
            )?;
            add_metadata_count(
                &mut count,
                field_info.data_references.len(),
                "FieldInfo.data_references",
            )?;
            add_metadata_count(
                &mut count,
                field_info.known_field_version.len(),
                "FieldInfo.known_field_version",
            )?;
            add_metadata_count(
                &mut count,
                usize::from(field_info.known_field_feature_identifier.is_some()),
                "FieldInfo.known_field_feature_identifier",
            )?;
        }
        if let Some(path) = &self.diff_field_path {
            add_metadata_count(
                &mut count,
                path.path.len(),
                "MessageInfo.diff_field_path.path",
            )?;
        }
        for path in &self.fields_to_remove {
            add_metadata_count(
                &mut count,
                path.path.len(),
                "MessageInfo.fields_to_remove.path",
            )?;
        }
        Ok(count)
    }
}

fn add_metadata_count(total: &mut usize, amount: usize, path: &str) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| overflow_error("metadata item", path))?;
    Ok(())
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
        Self::parse_with_limits(data, ArchiveLimits::default())
    }

    /// Parse a decompressed IWA stream under an explicit resource budget.
    pub fn parse_with_limits(data: &[u8], limits: ArchiveLimits) -> Result<Self> {
        let limits = limits.validate()?;
        if data.len() > limits.max_archive_bytes {
            return Err(limit_error(
                "archive byte",
                data.len(),
                limits.max_archive_bytes,
                "archive",
            ));
        }

        let mut objects = Vec::new();
        let mut cursor = Cursor::new(data);
        let mut total_messages = 0;
        let mut total_payload_bytes = 0;

        while cursor.position() < data.len() as u64 {
            if objects.len() >= limits.max_objects {
                return Err(limit_error(
                    "object",
                    objects.len() + 1,
                    limits.max_objects,
                    "archive",
                ));
            }
            let varint_start = cursor.position();
            let header_length = usize::try_from(varint::decode_varint(&mut cursor)?)
                .map_err(|_| Error::Archive("ArchiveInfo length exceeds usize".to_string()))?;
            if header_length > limits.max_header_bytes {
                return Err(limit_error(
                    "header byte",
                    header_length,
                    limits.max_header_bytes,
                    "ArchiveInfo",
                ));
            }
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
            archive_info.validate_with_limits(limits, "ArchiveInfo")?;
            checked_add_limited(
                &mut total_messages,
                archive_info.message_infos.len(),
                "message",
                limits.max_messages,
                "archive",
            )?;

            cursor.set_position(header_end as u64);
            let data_start = cursor.position();
            let object_prefix = usize::try_from(varint_length)
                .ok()
                .and_then(|length| length.checked_add(header_length))
                .ok_or_else(|| overflow_error("object byte", "ArchiveInfo"))?;
            if object_prefix > limits.max_object_bytes {
                return Err(limit_error(
                    "object byte",
                    object_prefix,
                    limits.max_object_bytes,
                    "ArchiveInfo",
                ));
            }
            let max_payload_bytes = limits.max_object_bytes - object_prefix;

            let mut payload_length = 0;
            for (message_index, info) in archive_info.message_infos.iter().enumerate() {
                let message_length = usize::try_from(info.length).map_err(|_| {
                    Error::Archive(format!(
                        "IWA message {message_index} length exceeds the platform usize limit"
                    ))
                })?;
                if message_length > limits.max_message_bytes {
                    return Err(limit_error(
                        "message byte",
                        message_length,
                        limits.max_message_bytes,
                        "MessageInfo.length",
                    ));
                }
                checked_add_limited(
                    &mut payload_length,
                    message_length,
                    "object payload byte",
                    max_payload_bytes,
                    "object",
                )?;
            }
            checked_add_limited(
                &mut total_payload_bytes,
                payload_length,
                "archive payload byte",
                limits.max_archive_bytes,
                "archive",
            )?;
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
                let message_length = usize::try_from(info.length).map_err(|_| {
                    Error::Archive("IWA message length exceeds the platform usize limit".to_owned())
                })?;
                let message_start = usize::try_from(cursor.position())
                    .map_err(|_| Error::Archive("IWA message offset exceeds usize".to_owned()))?;
                let message_end = message_start
                    .checked_add(message_length)
                    .ok_or_else(|| overflow_error("message byte", "message payload"))?;
                let message_data = data
                    .get(message_start..message_end)
                    .ok_or_else(|| Error::Archive("IWA message payload is truncated".to_owned()))?
                    .to_vec();
                cursor
                    .set_position(u64::try_from(message_end).map_err(|_| {
                        Error::Archive("IWA message offset exceeds u64".to_owned())
                    })?);
                messages.push(RawMessage {
                    type_: info.type_,
                    data: message_data,
                });
            }
            let data_length = cursor.position() - data_start;
            let decoded_messages = decode_raw_messages(&messages);
            let header_length_u64 = u64::try_from(header_length)
                .map_err(|_| Error::Archive("ArchiveInfo length exceeds u64".to_owned()))?;

            objects.push(ArchiveObject {
                archive_info,
                messages,
                decoded_messages,
                header_offset: varint_start,
                header_length: varint_length
                    .checked_add(header_length_u64)
                    .ok_or_else(|| overflow_error("header byte", "ArchiveInfo"))?,
                data_offset: data_start,
                data_length,
            });
        }

        let archive = Self { objects };
        archive.validate_with_limits(limits)?;
        Ok(archive)
    }

    /// Serialize this archive into the decompressed IWA stream representation.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_bytes_with_limits(ArchiveLimits::default())
    }

    /// Serialize this archive under an explicit resource budget.
    pub fn to_bytes_with_limits(&self, limits: ArchiveLimits) -> Result<Vec<u8>> {
        let limits = limits.validate()?;
        self.validate_with_limits(limits)?;
        let payload_bytes = self.objects.iter().try_fold(0usize, |total, object| {
            object.messages.iter().try_fold(total, |total, message| {
                total
                    .checked_add(message.data.len())
                    .ok_or_else(|| overflow_error("archive payload byte", "archive"))
            })
        })?;
        let object_overhead = self
            .objects
            .len()
            .checked_mul(16)
            .ok_or_else(|| overflow_error("archive byte", "serialization"))?;
        let capacity_hint = payload_bytes
            .checked_add(object_overhead)
            .ok_or_else(|| overflow_error("archive byte", "serialization"))?
            .min(limits.max_archive_bytes);
        let mut output = Vec::with_capacity(capacity_hint);

        for object in &self.objects {
            let mut info = object.archive_info.clone();
            for (message_info, message) in info.message_infos.iter_mut().zip(&object.messages) {
                message_info.type_ = message.type_;
                message_info.length = u32::try_from(message.data.len()).map_err(|_| {
                    Error::Archive("IWA message payload exceeds the u32 format limit".to_string())
                })?;
            }
            let encoded_info = protobuf::tsp::ArchiveInfo::from(&info);
            let header_length = encoded_info.encoded_len();
            if header_length > limits.max_header_bytes {
                return Err(limit_error(
                    "header byte",
                    header_length,
                    limits.max_header_bytes,
                    "serialization ArchiveInfo",
                ));
            }
            let header = encoded_info.encode_to_vec();
            let header_prefix = varint::encode_varint(
                u64::try_from(header.len())
                    .map_err(|_| Error::Archive("ArchiveInfo length exceeds u64".to_owned()))?,
            );
            let payload_length = object.messages.iter().try_fold(0usize, |total, message| {
                total
                    .checked_add(message.data.len())
                    .ok_or_else(|| overflow_error("object payload byte", "serialization"))
            })?;
            let object_length = header_prefix
                .len()
                .checked_add(header.len())
                .and_then(|length| length.checked_add(payload_length))
                .ok_or_else(|| overflow_error("object byte", "serialization"))?;
            if object_length > limits.max_object_bytes {
                return Err(limit_error(
                    "object byte",
                    object_length,
                    limits.max_object_bytes,
                    "serialization object",
                ));
            }
            let next_length = output
                .len()
                .checked_add(object_length)
                .ok_or_else(|| overflow_error("archive byte", "serialization"))?;
            if next_length > limits.max_archive_bytes {
                return Err(limit_error(
                    "archive byte",
                    next_length,
                    limits.max_archive_bytes,
                    "serialization archive",
                ));
            }
            output.reserve_exact(object_length);
            output.extend_from_slice(&header_prefix);
            output.extend_from_slice(&header);
            for message in &object.messages {
                output.extend_from_slice(&message.data);
            }
            debug_assert_eq!(output.len(), next_length, "object size mismatch");
        }
        Ok(output)
    }

    /// Validate object identifiers and message metadata before serialization.
    pub fn validate(&self) -> Result<()> {
        self.validate_with_limits(ArchiveLimits::default())
    }

    /// Validate archive structure and resource usage under an explicit budget.
    pub fn validate_with_limits(&self, limits: ArchiveLimits) -> Result<()> {
        validate_object_set(self.objects.iter(), limits.validate()?)
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
        self.insert_object_with_limits(object, ArchiveLimits::default())
    }

    /// Insert a new object under an explicit resource budget.
    pub fn insert_object_with_limits(
        &mut self,
        object: ArchiveObject,
        limits: ArchiveLimits,
    ) -> Result<()> {
        let limits = limits.validate()?;
        let identifier = object.archive_info.identifier.ok_or_else(|| {
            Error::Archive("Cannot insert an IWA object without an identifier".to_string())
        })?;
        if self.object(identifier).is_some() {
            return Err(Error::Archive(format!(
                "IWA object {identifier} already exists"
            )));
        }
        validate_object_set(std::iter::once(&object).chain(self.objects.iter()), limits)?;
        self.objects.push(object);
        Ok(())
    }

    /// Insert or replace an object, returning the previous value when present.
    pub fn upsert_object(&mut self, object: ArchiveObject) -> Result<Option<ArchiveObject>> {
        self.upsert_object_with_limits(object, ArchiveLimits::default())
    }

    /// Insert or replace an object under an explicit resource budget.
    pub fn upsert_object_with_limits(
        &mut self,
        object: ArchiveObject,
        limits: ArchiveLimits,
    ) -> Result<Option<ArchiveObject>> {
        let limits = limits.validate()?;
        let identifier = object.archive_info.identifier.ok_or_else(|| {
            Error::Archive("Cannot upsert an IWA object without an identifier".to_string())
        })?;
        if let Some(index) = self
            .objects
            .iter()
            .position(|item| item.archive_info.identifier == Some(identifier))
        {
            validate_object_set(
                self.objects.iter().enumerate().map(
                    |(current, item)| {
                        if current == index { &object } else { item }
                    },
                ),
                limits,
            )?;
            Ok(Some(std::mem::replace(&mut self.objects[index], object)))
        } else {
            validate_object_set(std::iter::once(&object).chain(self.objects.iter()), limits)?;
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

fn validate_object_set<'a, I>(objects: I, limits: ArchiveLimits) -> Result<()>
where
    I: IntoIterator<Item = &'a ArchiveObject>,
{
    let mut identifiers = HashSet::new();
    let mut object_count = 0usize;
    let mut message_count = 0usize;
    let mut payload_bytes = 0usize;

    for object in objects {
        object_count = object_count
            .checked_add(1)
            .ok_or_else(|| overflow_error("object", "archive"))?;
        if object_count > limits.max_objects {
            return Err(limit_error(
                "object",
                object_count,
                limits.max_objects,
                "archive",
            ));
        }
        let identifier = object.archive_info.identifier.ok_or_else(|| {
            Error::Archive("IWA object is missing its archive identifier".to_string())
        })?;
        if !identifiers.insert(identifier) {
            return Err(Error::Archive(format!(
                "Duplicate IWA object identifier {identifier}"
            )));
        }
        object.validate_with_limits(limits)?;
        checked_add_limited(
            &mut message_count,
            object.messages.len(),
            "message",
            limits.max_messages,
            "archive",
        )?;
        for message in &object.messages {
            checked_add_limited(
                &mut payload_bytes,
                message.data.len(),
                "archive payload byte",
                limits.max_archive_bytes,
                "archive",
            )?;
        }
    }
    Ok(())
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
        Self::new_with_limits(identifier, messages, ArchiveLimits::default())
    }

    /// Construct a new object under an explicit resource budget.
    pub fn new_with_limits(
        identifier: u64,
        messages: Vec<RawMessage>,
        limits: ArchiveLimits,
    ) -> Result<Self> {
        let limits = limits.validate()?;
        validate_raw_messages(&messages, limits, "new object")?;
        let mut message_infos = Vec::with_capacity(messages.len());
        for message in &messages {
            let length = u32::try_from(message.data.len()).map_err(|_| {
                Error::Archive("IWA message payload exceeds the u32 format limit".to_string())
            })?;
            message_infos.push(MessageInfo::new(message.type_, length));
        }
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
        self.validate_with_limits(ArchiveLimits::default())
    }

    pub fn validate_with_limits(&self, limits: ArchiveLimits) -> Result<()> {
        let limits = limits.validate()?;
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
        self.archive_info
            .validate_with_limits(limits, "object header")?;
        validate_raw_messages(&self.messages, limits, "object payload")?;
        for info in &self.archive_info.message_infos {
            let declared_length = usize::try_from(info.length).map_err(|_| {
                Error::Archive("IWA message length exceeds the platform usize limit".to_owned())
            })?;
            if declared_length > limits.max_message_bytes {
                return Err(limit_error(
                    "message byte",
                    declared_length,
                    limits.max_message_bytes,
                    "object MessageInfo.length",
                ));
            }
        }
        Ok(())
    }

    /// Replace a payload and keep its `MessageInfo` and decoded cache in sync.
    pub fn replace_message(&mut self, index: usize, message: RawMessage) -> Result<RawMessage> {
        self.replace_message_with_limits(index, message, ArchiveLimits::default())
    }

    /// Replace a payload under an explicit resource budget.
    pub fn replace_message_with_limits(
        &mut self,
        index: usize,
        message: RawMessage,
        limits: ArchiveLimits,
    ) -> Result<RawMessage> {
        let limits = limits.validate()?;
        self.validate_with_limits(limits)?;
        let old_length = self
            .messages
            .get(index)
            .ok_or_else(|| Error::Archive(format!("IWA message index {index} is out of bounds")))?
            .data
            .len();
        let new_length = message.data.len();
        check_message_length(new_length, limits, "replacement message")?;
        let payload_bytes = self.messages.iter().try_fold(0usize, |total, current| {
            total
                .checked_add(current.data.len())
                .ok_or_else(|| overflow_error("object payload byte", "replacement"))
        })?;
        let payload_without_old = payload_bytes
            .checked_sub(old_length)
            .ok_or_else(|| overflow_error("object payload byte", "replacement"))?;
        let next_payload_bytes = payload_without_old
            .checked_add(new_length)
            .ok_or_else(|| overflow_error("object payload byte", "replacement"))?;
        if next_payload_bytes > limits.max_object_bytes {
            return Err(limit_error(
                "object payload byte",
                next_payload_bytes,
                limits.max_object_bytes,
                "replacement object",
            ));
        }
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
        self.push_message_with_limits(message, ArchiveLimits::default())
    }

    pub fn push_message_with_limits(
        &mut self,
        message: RawMessage,
        limits: ArchiveLimits,
    ) -> Result<()> {
        let limits = limits.validate()?;
        self.validate_with_limits(limits)?;
        let next_count = self
            .messages
            .len()
            .checked_add(1)
            .ok_or_else(|| overflow_error("per-object message", "push message"))?;
        if next_count > limits.max_messages_per_object {
            return Err(limit_error(
                "per-object message",
                next_count,
                limits.max_messages_per_object,
                "push message",
            ));
        }
        let length = u32::try_from(message.data.len()).map_err(|_| {
            Error::Archive("IWA message payload exceeds the u32 format limit".to_string())
        })?;
        check_message_length(message.data.len(), limits, "push message")?;
        let payload_bytes =
            self.messages
                .iter()
                .try_fold(message.data.len(), |total, current| {
                    total
                        .checked_add(current.data.len())
                        .ok_or_else(|| overflow_error("object payload byte", "push message"))
                })?;
        if payload_bytes > limits.max_object_bytes {
            return Err(limit_error(
                "object payload byte",
                payload_bytes,
                limits.max_object_bytes,
                "push message",
            ));
        }
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

fn validate_raw_messages(messages: &[RawMessage], limits: ArchiveLimits, path: &str) -> Result<()> {
    if messages.len() > limits.max_messages_per_object {
        return Err(limit_error(
            "per-object message",
            messages.len(),
            limits.max_messages_per_object,
            path,
        ));
    }
    let mut payload_bytes = 0;
    for message in messages {
        check_message_length(message.data.len(), limits, path)?;
        checked_add_limited(
            &mut payload_bytes,
            message.data.len(),
            "object payload byte",
            limits.max_object_bytes,
            path,
        )?;
    }
    Ok(())
}

fn check_message_length(length: usize, limits: ArchiveLimits, path: &str) -> Result<()> {
    if length > limits.max_message_bytes {
        return Err(limit_error(
            "message byte",
            length,
            limits.max_message_bytes,
            path,
        ));
    }
    Ok(())
}

/// Raw protobuf payload data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawMessage {
    pub type_: u32,
    pub data: Vec<u8>,
}

fn decode_raw_messages(messages: &[RawMessage]) -> Vec<Box<dyn DecodedMessage>> {
    let mut decoded_messages =
        Vec::with_capacity(messages.len().min(ArchiveLimits::MAX_MESSAGES_PER_OBJECT));
    for message in messages.iter().take(ArchiveLimits::MAX_MESSAGES_PER_OBJECT) {
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
        bytes.extend(std::iter::repeat_n(0, 4096));
        bytes.push(0x80);

        let error = Archive::parse(&bytes).unwrap_err();

        assert!(matches!(
            error,
            Error::Archive(message)
                if message == "IWA object is missing its archive identifier"
        ));
    }

    fn encoded_object(identifier: u64, message_infos: Vec<protobuf::tsp::MessageInfo>) -> Vec<u8> {
        let header = protobuf::tsp::ArchiveInfo {
            identifier: Some(identifier),
            message_infos,
            should_merge: None,
        }
        .encode_to_vec();
        let mut bytes = varint::encode_varint(header.len() as u64);
        bytes.extend_from_slice(&header);
        bytes
    }

    #[test]
    fn archive_limits_are_finite_and_cross_field_safe() {
        let limits = ArchiveLimits::default();
        assert!(limits.max_archive_bytes() > 0);
        assert!(limits.max_messages() >= limits.max_messages_per_object());
        assert!(limits.max_archive_bytes() >= limits.max_object_bytes());
        assert!(limits.max_object_bytes() >= limits.max_message_bytes());
        assert!(ArchiveLimits::default().with_archive_bytes(0).is_err());
        assert!(ArchiveLimits::default().with_message_bytes(0).is_err());
        assert!(
            ArchiveLimits::default()
                .with_message_bytes(1)
                .is_ok_and(|limits| limits.max_message_bytes() == 1)
        );
    }

    #[test]
    fn rejects_aggregate_bytes_before_materializing_objects() {
        let bytes = encoded_object(1, Vec::new());
        let limits = ArchiveLimits::default()
            .with_archive_bytes(bytes.len() - 1)
            .unwrap();
        let error = Archive::parse_with_limits(&bytes, limits).unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("archive byte")));
    }

    #[test]
    fn rejects_object_and_message_count_limits_before_payload_work() {
        let bytes = [encoded_object(1, Vec::new()), encoded_object(2, Vec::new())].concat();
        let limits = ArchiveLimits::default().with_objects(1).unwrap();
        let error = Archive::parse_with_limits(&bytes, limits).unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("object limit")));

        let message_infos = vec![
            protobuf::tsp::MessageInfo {
                r#type: 1,
                length: 0,
                ..Default::default()
            },
            protobuf::tsp::MessageInfo {
                r#type: 2,
                length: 0,
                ..Default::default()
            },
        ];
        let bytes = encoded_object(1, message_infos);
        let limits = ArchiveLimits::default()
            .with_messages_per_object(1)
            .unwrap();
        let error = Archive::parse_with_limits(&bytes, limits).unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("per-object message")));
    }

    #[test]
    fn rejects_aggregate_message_and_payload_limits_fail_closed() {
        let message = || protobuf::tsp::MessageInfo {
            r#type: 1,
            length: 0,
            ..Default::default()
        };
        let bytes = [
            encoded_object(1, vec![message()]),
            encoded_object(2, vec![message()]),
        ]
        .concat();
        let limits = ArchiveLimits::default().with_messages(1).unwrap();
        let error = Archive::parse_with_limits(&bytes, limits).unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("message limit")));

        let bytes = encoded_object(
            1,
            vec![protobuf::tsp::MessageInfo {
                r#type: 7,
                length: 5,
                ..Default::default()
            }],
        );
        let limits = ArchiveLimits::default().with_message_bytes(4).unwrap();
        let error = Archive::parse_with_limits(&bytes, limits).unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("message byte")));
    }

    #[test]
    fn rejects_header_and_metadata_resource_limits() {
        let bytes = encoded_object(1, Vec::new());
        let limits = ArchiveLimits::default().with_header_bytes(1).unwrap();
        let error = Archive::parse_with_limits(&bytes, limits).unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("header byte")));

        let bytes = encoded_object(
            1,
            vec![protobuf::tsp::MessageInfo {
                r#type: 7,
                length: 0,
                object_references: vec![11, 12],
                ..Default::default()
            }],
        );
        let limits = ArchiveLimits::default().with_metadata_items(1).unwrap();
        let error = Archive::parse_with_limits(&bytes, limits).unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("metadata item")));
    }

    #[test]
    fn serialization_and_mutation_honor_limits_without_partial_state() {
        let archive = Archive {
            objects: vec![ArchiveObject::new(1, Vec::new()).unwrap()],
        };
        let limits = ArchiveLimits::default().with_header_bytes(1).unwrap();
        let error = archive.to_bytes_with_limits(limits).unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("header byte")));

        let mut object = ArchiveObject::new(1, Vec::new()).unwrap();
        let limits = ArchiveLimits::default().with_message_bytes(1).unwrap();
        let error = object
            .push_message_with_limits(
                RawMessage {
                    type_: 7,
                    data: vec![1, 2],
                },
                limits,
            )
            .unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("message byte")));
        assert!(object.messages.is_empty());
        assert!(object.archive_info.message_infos.is_empty());
    }

    #[test]
    fn bounded_header_reader_rejects_trailing_input() {
        let mut reader = Cursor::new(vec![0, 0]);
        let limits = ArchiveLimits::default().with_header_bytes(1).unwrap();
        let error = ArchiveInfo::parse_with_limits(&mut reader, limits).unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("header byte")));
    }
}
