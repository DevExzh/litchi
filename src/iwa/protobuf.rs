//! Protobuf Message Support for iWork IWA Files
//!
//! This module provides support for decoding Protocol Buffers messages
//! used in iWork IWA (iWork Archive) files using the prost crate.

use crate::iwa::{Error, Result};
use prost::Message;

// Include the generated protobuf definitions from build.rs
include!(concat!(env!("OUT_DIR"), "/iwa_protos.rs"));

/// A message decoder that can handle different iWork message types
pub struct MessageDecoder {
    /// Map of message type IDs to decoding functions
    decoders: std::collections::HashMap<u32, Box<dyn Fn(&[u8]) -> Result<Box<dyn DecodedMessage>>>>,
}

impl Default for MessageDecoder {
    fn default() -> Self {
        Self::new()
    }
}

impl MessageDecoder {
    /// Create a new message decoder
    pub fn new() -> Self {
        let mut decoder = MessageDecoder {
            decoders: std::collections::HashMap::new(),
        };

        // Register decoders for common message types
        decoder.register_decoder(1, |data| {
            let msg = tsp::ArchiveInfo::decode(data)?;
            Ok(Box::new(ArchiveInfoWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        decoder.register_decoder(2, |data| {
            let msg = tsp::MessageInfo::decode(data)?;
            Ok(Box::new(MessageInfoWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        // Register decoders for Pages document types
        decoder.register_decoder(1001, |data| {
            let msg = tp::DocumentArchive::decode(data)?;
            Ok(Box::new(PagesDocumentWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        // Register decoders for Numbers document types
        decoder.register_decoder(1002, |data| {
            let msg = tn::DocumentArchive::decode(data)?;
            Ok(Box::new(NumbersDocumentWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        decoder.register_decoder(1003, |data| {
            let msg = tn::SheetArchive::decode(data)?;
            Ok(Box::new(NumbersSheetWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        // Register decoders for Keynote document types
        decoder.register_decoder(1101, |data| {
            let msg = kn::ShowArchive::decode(data)?;
            Ok(Box::new(KeynoteShowWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        decoder.register_decoder(1102, |data| {
            let msg = kn::SlideArchive::decode(data)?;
            Ok(Box::new(KeynoteSlideWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        // Register decoders for text content
        decoder.register_decoder(200, |data| {
            let msg = tswp::StorageArchive::decode(data)?;
            Ok(Box::new(StorageArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        // Register decoders for table content (Numbers)
        decoder.register_decoder(100, |data| {
            let msg = tst::TableModelArchive::decode(data)?;
            Ok(Box::new(TableModelWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        // Register additional StorageArchive decoders for common text-containing message types
        // These are commonly found in iWork files
        decoder.register_decoder(201, |data| {
            let msg = tswp::StorageArchive::decode(data)?;
            Ok(Box::new(StorageArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
        });
        decoder.register_decoder(202, |data| {
            let msg = tswp::StorageArchive::decode(data)?;
            Ok(Box::new(StorageArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
        });
        decoder.register_decoder(203, |data| {
            let msg = tswp::StorageArchive::decode(data)?;
            Ok(Box::new(StorageArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
        });
        decoder.register_decoder(204, |data| {
            let msg = tswp::StorageArchive::decode(data)?;
            Ok(Box::new(StorageArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
        });
        decoder.register_decoder(205, |data| {
            let msg = tswp::StorageArchive::decode(data)?;
            Ok(Box::new(StorageArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        // Add decoder for message type 2022 which is very common in test files
        decoder.register_decoder(2022, |data| {
            let msg = tswp::StorageArchive::decode(data)?;
            Ok(Box::new(StorageArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
        });

        decoder
    }

    /// Register a decoder for a specific message type
    pub fn register_decoder<F>(&mut self, message_type: u32, decoder_fn: F)
    where
        F: Fn(&[u8]) -> Result<Box<dyn DecodedMessage>> + 'static,
    {
        self.decoders.insert(message_type, Box::new(decoder_fn));
    }

    /// Decode a message of the given type
    pub fn decode(&self, message_type: u32, data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
        if let Some(decoder) = self.decoders.get(&message_type) {
            decoder(data)
        } else {
            Err(Error::UnsupportedMessageType(message_type))
        }
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
        if let Some(ref name) = self.0.name {
            if !name.is_empty() {
                text.push(name.clone());
            }
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
        let decoder = MessageDecoder::new();
        assert!(decoder.decoders.contains_key(&1)); // ArchiveInfo
        assert!(decoder.decoders.contains_key(&2)); // MessageInfo
        assert!(decoder.decoders.contains_key(&100)); // TableModelArchive
        assert!(decoder.decoders.contains_key(&200)); // StorageArchive
        assert!(decoder.decoders.contains_key(&201)); // StorageArchive variant
        assert!(decoder.decoders.contains_key(&202)); // StorageArchive variant
        assert!(decoder.decoders.contains_key(&2022)); // Common StorageArchive type
        assert!(decoder.decoders.contains_key(&1001)); // Pages Document
        assert!(decoder.decoders.contains_key(&1002)); // Numbers Document
        assert!(decoder.decoders.contains_key(&1003)); // Numbers Sheet
        assert!(decoder.decoders.contains_key(&1101)); // Keynote Show
        assert!(decoder.decoders.contains_key(&1102)); // Keynote Slide
    }

    #[test]
    fn test_unsupported_message_type() {
        let decoder = MessageDecoder::new();
        let result = decoder.decode(999, &[]);
        assert!(matches!(result, Err(Error::UnsupportedMessageType(999))));
    }
}

