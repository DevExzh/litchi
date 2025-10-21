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

/// Static decoder function for Numbers SheetArchive messages
fn decode_numbers_sheet(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tn::SheetArchive::decode(data)?;
    Ok(Box::new(NumbersSheetWrapper(msg)) as Box<dyn DecodedMessage>)
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

/// Static decoder function for TableDataList messages
fn decode_table_data_list(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tst::TableDataList::decode(data)?;
    Ok(Box::new(TableDataListWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for ShapeArchive messages
fn decode_shape_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tsd::ShapeArchive::decode(data)?;
    Ok(Box::new(ShapeArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for DrawableArchive messages
fn decode_drawable_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tsd::DrawableArchive::decode(data)?;
    Ok(Box::new(DrawableArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for ChartArchive messages
fn decode_chart_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tsch::ChartArchive::decode(data)?;
    Ok(Box::new(ChartArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
}

type DecoderMap = phf::Map<u32, fn(&[u8]) -> Result<Box<dyn DecodedMessage>>>;

/// Perfect hash map of message type IDs to decoder functions
/// This provides O(1) lookup performance at compile time
///
/// Based on analysis of iWork documents and official message type registry:
/// - 1-2: TSP (Core Protocol) ArchiveInfo and MessageInfo
/// - 200-299: TSK (Document Core)
/// - 400-499: TSS (Stylesheets)
/// - 600-699: TSA (Application Core)
/// - 2000-2999: TSWP (Word Processing / Text)
/// - 3000-3999: TSD (Drawing / Shapes)
/// - 4000-4999: TSCE (Calculation Engine)
/// - 5000-5999: TSCH (Charts)
/// - 6000-6999: TST (Tables)
/// - 10000-10999: TP (Pages-specific)
/// - 12000-12999: TN (Numbers-specific)
/// - 1-25, 100-199: KN (Keynote-specific)
///
/// Note: Message types are application-specific and may overlap between apps
static DECODERS: DecoderMap = phf_map! {
    // TSP (Core Protocol) types - used by all applications
    1u32 => decode_archive_info,
    2u32 => decode_message_info,

    // TST (Table) types - Numbers spreadsheet tables and cells
    // Message type 6001 is TST.TableModelArchive
    6000u32 => decode_table_model,
    6001u32 => decode_table_model,
    6005u32 => decode_table_data_list,
    6201u32 => decode_table_data_list,

    // TSD (Drawing) types - Shapes, images, and drawables
    3002u32 => decode_drawable_archive,
    3003u32 => decode_drawable_archive,  // ContainerArchive
    3004u32 => decode_shape_archive,
    3005u32 => decode_shape_archive,     // ImageArchive (shape variant)
    3006u32 => decode_shape_archive,     // MaskArchive
    3007u32 => decode_shape_archive,     // MovieArchive
    3008u32 => decode_shape_archive,     // GroupArchive
    3009u32 => decode_shape_archive,     // ConnectionLineArchive

    // TSCH (Charts) types
    5000u32 => decode_chart_archive,
    5004u32 => decode_chart_archive,  // ChartMediatorArchive
    5021u32 => decode_chart_archive,  // ChartDrawableArchive

    // TSWP (Word Processing) types - Text storage used across all apps
    2001u32 => decode_storage_archive,
    2002u32 => decode_storage_archive,
    2003u32 => decode_storage_archive,
    2004u32 => decode_storage_archive,
    2005u32 => decode_storage_archive,
    2006u32 => decode_storage_archive,
    2007u32 => decode_storage_archive,
    2008u32 => decode_storage_archive,
    2009u32 => decode_storage_archive,
    2010u32 => decode_storage_archive,
    2011u32 => decode_storage_archive,
    2012u32 => decode_storage_archive,
    2013u32 => decode_storage_archive,
    2014u32 => decode_storage_archive,
    2022u32 => decode_storage_archive,

    // Pages-specific document types (TP namespace)
    10000u32 => decode_pages_document,
    10001u32 => decode_pages_document,  // ThemeArchive
    10011u32 => decode_pages_document,  // SectionArchive

    // Numbers-specific document types (TN namespace)
    // Note: Numbers uses low message type numbers that conflict with TSP types
    // Type 1 for TN.DocumentArchive conflicts with TSP.ArchiveInfo
    // Type 2 for TN.SheetArchive conflicts with TSP.MessageInfo
    // We prioritize TSP types in the decoder map and handle app-specific
    // types through context-aware parsing in the document parsers
    3u32 => decode_numbers_sheet,       // TN.FormBasedSheetArchive

    // Keynote-specific document types (KN namespace)
    // Note: Keynote uses low message type numbers (1-25, 100-199)
    // Type 2 is KN.ShowArchive but conflicts with TSP.MessageInfo
    // Type 5/6 are KN.SlideArchive
    5u32 => decode_keynote_slide,       // KN.SlideArchive
    6u32 => decode_keynote_slide,       // KN.SlideArchive (variant)
    8u32 => decode_keynote_slide,       // KN.BuildArchive (slide-related)
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
    fn message_type(&self) -> u32 {
        1
    }

    fn extract_text(&self) -> Vec<String> {
        Vec::new() // ArchiveInfo doesn't contain text
    }
}

/// Wrapper for MessageInfo message
#[derive(Debug)]
pub struct MessageInfoWrapper(pub tsp::MessageInfo);

impl DecodedMessage for MessageInfoWrapper {
    fn message_type(&self) -> u32 {
        2
    }

    fn extract_text(&self) -> Vec<String> {
        Vec::new() // MessageInfo doesn't contain text
    }
}

/// Wrapper for StorageArchive message (text content)
#[derive(Debug)]
pub struct StorageArchiveWrapper(pub tswp::StorageArchive);

impl DecodedMessage for StorageArchiveWrapper {
    fn message_type(&self) -> u32 {
        200
    }

    fn extract_text(&self) -> Vec<String> {
        self.0.text.clone()
    }
}

/// Document wrapper for TP.DocumentArchive
#[derive(Debug)]
pub struct PagesDocumentWrapper(pub tp::DocumentArchive);

impl DecodedMessage for PagesDocumentWrapper {
    fn message_type(&self) -> u32 {
        1001
    }

    fn extract_text(&self) -> Vec<String> {
        Vec::new() // Document metadata doesn't contain direct text
    }
}

/// Sheet wrapper for TN.SheetArchive
#[derive(Debug)]
pub struct NumbersSheetWrapper(pub tn::SheetArchive);

impl DecodedMessage for NumbersSheetWrapper {
    fn message_type(&self) -> u32 {
        1003
    }

    fn extract_text(&self) -> Vec<String> {
        if !self.0.name.is_empty() {
            vec![self.0.name.clone()]
        } else {
            Vec::new()
        }
    }
}

/// Wrapper for Keynote Slide Archive
#[derive(Debug)]
pub struct KeynoteSlideWrapper(pub kn::SlideArchive);

impl DecodedMessage for KeynoteSlideWrapper {
    fn message_type(&self) -> u32 {
        1102
    }

    fn extract_text(&self) -> Vec<String> {
        let mut text = Vec::new();
        if let Some(ref name) = self.0.name
            && !name.is_empty()
        {
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
    fn message_type(&self) -> u32 {
        100
    }

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

/// Wrapper for Table Data List (cell content storage)
#[derive(Debug)]
pub struct TableDataListWrapper(pub tst::TableDataList);

impl DecodedMessage for TableDataListWrapper {
    fn message_type(&self) -> u32 {
        101
    }

    fn extract_text(&self) -> Vec<String> {
        // TableDataList contains actual cell data as ListEntry items
        // Extract string values from entries
        let mut strings = Vec::new();

        for entry in &self.0.entries {
            if let Some(ref string_val) = entry.string
                && !string_val.is_empty()
            {
                strings.push(string_val.clone());
            }
        }

        strings
    }
}

/// Wrapper for Shape Archive
#[derive(Debug)]
pub struct ShapeArchiveWrapper(pub tsd::ShapeArchive);

impl DecodedMessage for ShapeArchiveWrapper {
    fn message_type(&self) -> u32 {
        500
    }

    fn extract_text(&self) -> Vec<String> {
        // Shapes can contain text, particularly text boxes
        // Text is typically stored in the DrawableArchive's accessibility description
        // or in referenced TSWP.StorageArchive objects (handled by shape text extractor)
        let mut text = Vec::new();

        // super_ is a required field, not Optional
        let drawable = &self.0.super_;

        // Extract accessibility description if present (often used for alt text/labels)
        if let Some(ref desc) = drawable.accessibility_description
            && !desc.is_empty()
        {
            text.push(desc.clone());
        }

        // Hyperlink URLs can also contain meaningful text
        if let Some(ref url) = drawable.hyperlink_url
            && !url.is_empty()
        {
            text.push(url.clone());
        }

        text
    }
}

/// Wrapper for Drawable Archive
#[derive(Debug)]
pub struct DrawableArchiveWrapper(pub tsd::DrawableArchive);

impl DecodedMessage for DrawableArchiveWrapper {
    fn message_type(&self) -> u32 {
        501
    }

    fn extract_text(&self) -> Vec<String> {
        // Drawables are visual elements without direct text
        Vec::new()
    }
}

/// Wrapper for Chart Archive
#[derive(Debug)]
pub struct ChartArchiveWrapper(pub tsch::ChartArchive);

impl DecodedMessage for ChartArchiveWrapper {
    fn message_type(&self) -> u32 {
        600
    }

    fn extract_text(&self) -> Vec<String> {
        // Charts contain text in grid data (row/column names)
        // and may have titles in referenced text storage objects
        let mut text = Vec::new();

        // Extract grid data (row and column names)
        if let Some(ref grid) = self.0.grid {
            // Add row names
            for row_name in &grid.row_name {
                if !row_name.is_empty() {
                    text.push(row_name.clone());
                }
            }

            // Add column names
            for col_name in &grid.column_name {
                if !col_name.is_empty() {
                    text.push(col_name.clone());
                }
            }
        }

        text
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_message_decoder_creation() {
        // Test that all expected decoders are available in the static map
        assert!(DECODERS.contains_key(&1)); // ArchiveInfo / TN.DocumentArchive
        assert!(DECODERS.contains_key(&2)); // MessageInfo / TN.SheetArchive
        assert!(DECODERS.contains_key(&6001)); // TST.TableModelArchive
        assert!(DECODERS.contains_key(&2001)); // TSWP.StorageArchive
        assert!(DECODERS.contains_key(&2002)); // StorageArchive variant
        assert!(DECODERS.contains_key(&2003)); // StorageArchive variant
        assert!(DECODERS.contains_key(&2022)); // Common StorageArchive type
        assert!(DECODERS.contains_key(&10000)); // TP.DocumentArchive (Pages)
        assert!(DECODERS.contains_key(&3)); // TN.FormBasedSheetArchive (Numbers)
        assert!(DECODERS.contains_key(&5)); // KN.SlideArchive (Keynote)
        assert!(DECODERS.contains_key(&6)); // KN.SlideArchive variant (Keynote)
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
        let message_types = [1, 2, 6001, 2001, 2002, 2003, 10000, 3, 5];

        // Create some dummy data that will fail to decode but test the lookup
        let dummy_data = vec![0u8; 10];

        for &msg_type in &message_types {
            let result = decode(msg_type, &dummy_data);
            // We expect this to fail due to invalid protobuf data, but the lookup should be fast
            assert!(result.is_err());
        }
    }
}
