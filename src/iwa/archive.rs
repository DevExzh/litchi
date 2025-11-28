//! IWA Archive Format Parser
//!
//! This module handles parsing of IWA (iWork Archive) files, which contain
//! Protocol Buffers-encoded messages with ArchiveInfo and MessageInfo headers.

use std::io::Read;

use crate::iwa::protobuf::{DecodedMessage, decode};
use crate::iwa::varint;
use crate::iwa::{Error, Result};
use prost::Message;

/// Archive information header for each object in an IWA file
#[derive(Debug, Clone, PartialEq)]
pub struct ArchiveInfo {
    /// Unique identifier for this archive across the document
    pub identifier: Option<u64>,
    /// Information about the messages contained in this archive
    pub message_infos: Vec<MessageInfo>,
}

impl ArchiveInfo {
    /// Parse ArchiveInfo from a reader
    pub fn parse<R: Read>(reader: &mut R) -> Result<Self> {
        let mut identifier = None;
        let mut message_infos = Vec::new();

        // Parse Protocol Buffer fields
        while let Ok((field_number, wire_type)) = Self::read_field_header(reader) {
            match (field_number, wire_type) {
                (1, 0) => {
                    // identifier (varint)
                    identifier = Some(varint::decode_varint(reader)?);
                },
                (2, 2) => {
                    // message_infos (length-delimited, repeated)
                    let length = varint::decode_varint(reader)?;
                    let mut data = vec![0u8; length as usize];
                    reader.read_exact(&mut data)?;
                    let mut cursor = std::io::Cursor::new(data);
                    message_infos.push(MessageInfo::parse(&mut cursor)?);
                },
                _ => {
                    // Skip unknown fields
                    Self::skip_field(reader, wire_type)?;
                },
            }
        }

        Ok(ArchiveInfo {
            identifier,
            message_infos,
        })
    }

    fn read_field_header<R: Read>(reader: &mut R) -> Result<(u32, u32)> {
        let tag = varint::decode_varint(reader)?;
        let field_number = (tag >> 3) as u32;
        let wire_type = (tag & 0x07) as u32;
        Ok((field_number, wire_type))
    }

    fn skip_field<R: Read>(reader: &mut R, wire_type: u32) -> Result<()> {
        match wire_type {
            0 => {
                // varint
                varint::decode_varint(reader)?;
            },
            1 => {
                // 64-bit
                let mut buf = [0u8; 8];
                reader.read_exact(&mut buf)?;
            },
            2 => {
                // length-delimited
                let length = varint::decode_varint(reader)?;
                let mut buf = vec![0u8; length as usize];
                reader.read_exact(&mut buf)?;
            },
            5 => {
                // 32-bit
                let mut buf = [0u8; 4];
                reader.read_exact(&mut buf)?;
            },
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "Unknown wire type: {}",
                    wire_type
                )));
            },
        }
        Ok(())
    }
}

/// Information about a specific message within an archive
#[derive(Debug, Clone, PartialEq)]
pub struct MessageInfo {
    /// Message type identifier (maps to specific protobuf message types)
    pub type_: u32,
    /// Version information for the message format
    pub versions: Vec<u32>,
    /// Length of the message data in bytes
    pub length: u32,
}

impl MessageInfo {
    /// Parse MessageInfo from a reader
    pub fn parse<R: Read>(reader: &mut R) -> Result<Self> {
        let mut type_ = 0;
        let mut versions = Vec::new();
        let mut length = 0;

        while let Ok((field_number, wire_type)) = Self::read_field_header(reader) {
            match (field_number, wire_type) {
                (1, 0) => {
                    // type (varint)
                    type_ = varint::decode_varint(reader)? as u32;
                },
                (2, 0) => {
                    // version (varint, packed repeated)
                    versions.push(varint::decode_varint(reader)? as u32);
                },
                (3, 0) => {
                    // length (varint)
                    length = varint::decode_varint(reader)? as u32;
                },
                _ => {
                    // Skip unknown fields
                    Self::skip_field(reader, wire_type)?;
                },
            }
        }

        Ok(MessageInfo {
            type_,
            versions,
            length,
        })
    }

    fn read_field_header<R: Read>(reader: &mut R) -> Result<(u32, u32)> {
        ArchiveInfo::read_field_header(reader)
    }

    fn skip_field<R: Read>(reader: &mut R, wire_type: u32) -> Result<()> {
        ArchiveInfo::skip_field(reader, wire_type)
    }
}

/// A parsed IWA archive containing multiple objects
#[derive(Debug)]
pub struct Archive {
    /// The objects contained in this archive
    pub objects: Vec<ArchiveObject>,
}

impl Archive {
    /// Parse an IWA archive from decompressed data
    ///
    /// This function tracks byte offsets for each object to enable efficient
    /// lazy loading and partial parsing. The implementation follows the IWA
    /// format specification from Apple's iWorkFileFormat documentation.
    ///
    /// # Performance
    ///
    /// O(n) where n is the number of bytes in the decompressed data.
    /// Memory usage is proportional to the number of objects.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut objects = Vec::new();
        let mut cursor = std::io::Cursor::new(data);

        while cursor.position() < data.len() as u64 {
            // Track the start of this object's header (before the varint length)
            let varint_start_pos = cursor.position();

            // Read archive info length
            let archive_info_length = varint::decode_varint(&mut cursor)? as usize;

            // The header starts after the varint that encodes its length
            let header_start_pos = cursor.position();
            let varint_length = header_start_pos - varint_start_pos;

            // Read archive info
            let mut archive_info_data = vec![0u8; archive_info_length];
            cursor.read_exact(&mut archive_info_data)?;
            let mut archive_info_cursor = std::io::Cursor::new(archive_info_data);
            let archive_info = ArchiveInfo::parse(&mut archive_info_cursor)?;

            // Calculate total data length from all message infos
            let total_data_length: u64 = archive_info
                .message_infos
                .iter()
                .map(|mi| mi.length as u64)
                .sum();

            // Data starts immediately after the header
            let data_start_pos = cursor.position();

            // Read message data
            let mut messages = Vec::new();
            let mut decoded_messages = Vec::new();

            for message_info in &archive_info.message_infos {
                let mut message_data = vec![0u8; message_info.length as usize];
                cursor.read_exact(&mut message_data)?;

                let raw_message = RawMessage {
                    type_: message_info.type_,
                    data: message_data.clone(),
                };

                messages.push(raw_message);

                // Try to decode the message using prost
                match decode(message_info.type_, &message_data) {
                    Ok(decoded) => decoded_messages.push(decoded),
                    Err(_) => {
                        // Message type not registered - try parsing as StorageArchive anyway
                        // since many message types might contain text
                        if let Ok(storage_msg) =
                            crate::iwa::protobuf::tswp::StorageArchive::decode(&*message_data)
                        {
                            let wrapper = crate::iwa::protobuf::StorageArchiveWrapper(storage_msg);
                            decoded_messages
                                .push(Box::new(wrapper)
                                    as Box<dyn crate::iwa::protobuf::DecodedMessage>);
                        }
                    },
                }
            }

            objects.push(ArchiveObject {
                archive_info,
                messages,
                decoded_messages,
                header_offset: varint_start_pos,
                header_length: varint_length + archive_info_length as u64,
                data_offset: data_start_pos,
                data_length: total_data_length,
            });
        }

        Ok(Archive { objects })
    }
}

/// A single object within an IWA archive
#[derive(Debug)]
pub struct ArchiveObject {
    /// Archive metadata
    pub archive_info: ArchiveInfo,
    /// Raw message data (protobuf-encoded)
    pub messages: Vec<RawMessage>,
    /// Decoded message objects (if successfully decoded)
    pub decoded_messages: Vec<Box<dyn DecodedMessage>>,
    /// Byte offset of the ArchiveInfo header in the decompressed stream
    pub header_offset: u64,
    /// Length of the ArchiveInfo header in bytes
    pub header_length: u64,
    /// Byte offset of the message data (after the ArchiveInfo header)
    pub data_offset: u64,
    /// Total length of all message data in bytes
    pub data_length: u64,
}

/// Raw protobuf message data
#[derive(Debug, Clone)]
pub struct RawMessage {
    /// Message type identifier
    pub type_: u32,
    /// Raw protobuf data
    pub data: Vec<u8>,
}

impl ArchiveObject {
    /// Extract all text content from decoded messages
    pub fn extract_text(&self) -> Vec<String> {
        let mut all_text = Vec::new();
        for decoded_msg in &self.decoded_messages {
            all_text.extend(decoded_msg.extract_text());
        }
        all_text
    }

    /// Get the primary message type from this object
    pub fn primary_message_type(&self) -> Option<u32> {
        self.messages.first().map(|msg| msg.type_)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_archive_info_parsing() {
        // Create a minimal ArchiveInfo protobuf message
        // Field 1 (identifier): varint 123
        // Field 2 (message_infos): length-delimited with a MessageInfo
        let mut data = Vec::new();

        // Field 1: identifier = 123
        data.extend(varint::encode_varint(1 << 3)); // tag: field 1, wire type 0
        data.extend(varint::encode_varint(123));

        // Field 2: message_infos (simplified)
        data.extend(varint::encode_varint((2 << 3) | 2)); // tag: field 2, wire type 2
        let message_info_data = vec![
            0x08, 0x01, // type = 1
            0x18, 0x05, // length = 5
        ];
        data.extend(varint::encode_varint(message_info_data.len() as u64));
        data.extend(message_info_data);

        let mut cursor = std::io::Cursor::new(data);
        let archive_info = ArchiveInfo::parse(&mut cursor).unwrap();

        assert_eq!(archive_info.identifier, Some(123));
        assert_eq!(archive_info.message_infos.len(), 1);
        assert_eq!(archive_info.message_infos[0].type_, 1);
        assert_eq!(archive_info.message_infos[0].length, 5);
    }

    #[test]
    fn test_message_info_parsing() {
        // Create a MessageInfo protobuf message
        let data = vec![
            0x08, 0x2A, // type = 42
            0x10, 0x01, // version = 1
            0x18, 0x0A, // length = 10
        ];

        let mut cursor = std::io::Cursor::new(data);
        let message_info = MessageInfo::parse(&mut cursor).unwrap();

        assert_eq!(message_info.type_, 42);
        assert_eq!(message_info.versions, vec![1]);
        assert_eq!(message_info.length, 10);
    }
}
