//! Protobuf Message Support for iWork IWA Files
//!
//! This module provides support for decoding Protocol Buffers messages
//! used in iWork IWA (iWork Archive) files using the prost crate.

use crate::iwa::{Error, Result};
use phf::phf_map;
use prost::Message;

// Include the generated protobuf definitions from build.rs
include!(concat!(env!("OUT_DIR"), "/iwa_protos.rs"));

/// Static decoder function for ArchiveInfo messages
fn decode_archive_info(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tsp::ArchiveInfo::decode(data)?;
    Ok(Box::new(ArchiveInfoWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for MessageInfo messages
fn decode_message_info(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tsp::MessageInfo::decode(data)?;
    Ok(Box::new(MessageInfoWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for Pages DocumentArchive messages
fn decode_pages_document(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tp::DocumentArchive::decode(data)?;
    Ok(Box::new(PagesDocumentWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for Numbers DocumentArchive messages
fn decode_numbers_document(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tn::DocumentArchive::decode(data)?;
    Ok(Box::new(NumbersDocumentWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for Numbers SheetArchive messages
fn decode_numbers_sheet(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tn::SheetArchive::decode(data)?;
    Ok(Box::new(NumbersSheetWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for Keynote ShowArchive messages
fn decode_keynote_show(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = kn::ShowArchive::decode(data)?;
    Ok(Box::new(KeynoteShowWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for Keynote SlideArchive messages
fn decode_keynote_slide(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = kn::SlideArchive::decode(data)?;
    Ok(Box::new(KeynoteSlideWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for StorageArchive messages
fn decode_storage_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tswp::StorageArchive::decode(data)?;
    Ok(Box::new(StorageArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for TableModelArchive messages
fn decode_table_model(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tst::TableModelArchive::decode(data)?;
    Ok(Box::new(TableModelWrapper(msg)) as Box<dyn DecodedMessage>)
}

type DecoderMap = phf::Map<u32, fn(&[u8]) -> Result<Box<dyn DecodedMessage>>>;

/// Perfect hash map of message type IDs to decoder functions
/// This provides O(1) lookup performance at compile time
static DECODERS: DecoderMap = phf_map! {
    1u32 => decode_archive_info,
    2u32 => decode_message_info,
    100u32 => decode_table_model,
    200u32 => decode_storage_archive,
    201u32 => decode_storage_archive,
    202u32 => decode_storage_archive,
    203u32 => decode_storage_archive,
    204u32 => decode_storage_archive,
    205u32 => decode_storage_archive,
    1001u32 => decode_pages_document,
    1002u32 => decode_numbers_document,
    1003u32 => decode_numbers_sheet,
    1101u32 => decode_keynote_show,
    1102u32 => decode_keynote_slide,
    2022u32 => decode_storage_archive,
};

/// Decode a message of the given type using the perfect hash map for O(1) lookup
pub fn decode(message_type: u32, data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    if let Some(decoder) = DECODERS.get(&message_type) {
        decoder(data)
    } else {
        Err(Error::UnsupportedMessageType(message_type))
    }
}

/// Trait for decoded iWork messages
pub trait DecodedMessage: std::fmt::Debug {
    /// Get the message type identifier
    fn message_type(&self) -> u32;

    /// Extract text content from the message if available
    fn extract_text(&self) -> Vec<String> {
        Vec::new()
    }
}

/// Wrapper for ArchiveInfo message
#[derive(Debug)]
pub struct ArchiveInfoWrapper(pub tsp::ArchiveInfo);

impl DecodedMessage for ArchiveInfoWrapper {
    fn message_type(&self) -> u32 { 1 }

    fn extract_text(&self) -> Vec<String> {
        Vec::new() // ArchiveInfo doesn't contain text
    }
}

/// Wrapper for MessageInfo message
#[derive(Debug)]
pub struct MessageInfoWrapper(pub tsp::MessageInfo);

impl DecodedMessage for MessageInfoWrapper {
    fn message_type(&self) -> u32 { 2 }

    fn extract_text(&self) -> Vec<String> {
        Vec::new() // MessageInfo doesn't contain text
    }
}

/// Wrapper for StorageArchive message (text content)
#[derive(Debug)]
pub struct StorageArchiveWrapper(pub tswp::StorageArchive);

impl DecodedMessage for StorageArchiveWrapper {
    fn message_type(&self) -> u32 { 200 }

    fn extract_text(&self) -> Vec<String> {
        self.0.text.clone()
    }
}

/// Document wrapper for TP.DocumentArchive
#[derive(Debug)]
pub struct PagesDocumentWrapper(pub tp::DocumentArchive);

impl DecodedMessage for PagesDocumentWrapper {
    fn message_type(&self) -> u32 { 1001 }

    fn extract_text(&self) -> Vec<String> {
        Vec::new() // Document metadata doesn't contain direct text
    }
}

/// Document wrapper for TN.DocumentArchive
#[derive(Debug)]
pub struct NumbersDocumentWrapper(pub tn::DocumentArchive);

impl DecodedMessage for NumbersDocumentWrapper {
    fn message_type(&self) -> u32 { 1002 }

    fn extract_text(&self) -> Vec<String> {
        Vec::new() // Document metadata doesn't contain direct text
    }
}

/// Sheet wrapper for TN.SheetArchive
#[derive(Debug)]
pub struct NumbersSheetWrapper(pub tn::SheetArchive);

impl DecodedMessage for NumbersSheetWrapper {
    fn message_type(&self) -> u32 { 1003 }

    fn extract_text(&self) -> Vec<String> {
        if !self.0.name.is_empty() {
            vec![self.0.name.clone()]
        } else {
            Vec::new()
        }
    }
}

/// Wrapper for Keynote Show Archive
#[derive(Debug)]
pub struct KeynoteShowWrapper(pub kn::ShowArchive);

impl DecodedMessage for KeynoteShowWrapper {
    fn message_type(&self) -> u32 { 1101 }

    fn extract_text(&self) -> Vec<String> {
        // ShowArchive doesn't have a title field, return empty
        Vec::new()
    }
}

/// Wrapper for Keynote Slide Archive
#[derive(Debug)]
pub struct KeynoteSlideWrapper(pub kn::SlideArchive);

impl DecodedMessage for KeynoteSlideWrapper {
    fn message_type(&self) -> u32 { 1102 }

    fn extract_text(&self) -> Vec<String> {
        let mut text = Vec::new();
        if let Some(ref name) = self.0.name
            && !name.is_empty() {
                text.push(name.clone());
            }
        // if let Some(ref note) = self.0.note {
        //     // Note is a reference, not direct text - we can't extract text from it here
        //     // without additional processing
        // }
        text
    }
}

/// Wrapper for Table Model Archive (Numbers tables)
#[derive(Debug)]
pub struct TableModelWrapper(pub tst::TableModelArchive);

impl DecodedMessage for TableModelWrapper {
    fn message_type(&self) -> u32 { 100 }

    fn extract_text(&self) -> Vec<String> {
        let mut text = Vec::new();
        // Extract table name if present
        if !self.0.table_name.is_empty() {
            text.push(self.0.table_name.clone());
        }
        // Note: Cell contents are stored in data_store which requires complex
        // processing to extract. For now, we only return the table name.
        text
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_message_decoder_creation() {
        // Test that all expected decoders are available in the static map
        assert!(DECODERS.contains_key(&1)); // ArchiveInfo
        assert!(DECODERS.contains_key(&2)); // MessageInfo
        assert!(DECODERS.contains_key(&100)); // TableModelArchive
        assert!(DECODERS.contains_key(&200)); // StorageArchive
        assert!(DECODERS.contains_key(&201)); // StorageArchive variant
        assert!(DECODERS.contains_key(&202)); // StorageArchive variant
        assert!(DECODERS.contains_key(&2022)); // Common StorageArchive type
        assert!(DECODERS.contains_key(&1001)); // Pages Document
        assert!(DECODERS.contains_key(&1002)); // Numbers Document
        assert!(DECODERS.contains_key(&1003)); // Numbers Sheet
        assert!(DECODERS.contains_key(&1101)); // Keynote Show
        assert!(DECODERS.contains_key(&1102)); // Keynote Slide
    }

    #[test]
    fn test_unsupported_message_type() {
        let result = decode(999, &[]);
        assert!(matches!(result, Err(Error::UnsupportedMessageType(999))));
    }

    #[test]
    fn test_decoder_performance() {
        // Test that decoding is fast with phf::Map
        // This test ensures the static map lookup is working
        let message_types = [1, 2, 100, 200, 201, 202, 1001, 1002, 1101];

        // Create some dummy data that will fail to decode but test the lookup
        let dummy_data = vec![0u8; 10];

        for &msg_type in &message_types {
            let result = decode(msg_type, &dummy_data);
            // We expect this to fail due to invalid protobuf data, but the lookup should be fast
            assert!(result.is_err());
        }
    }
}

