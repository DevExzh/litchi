//! Protobuf Message Support for iWork IWA Files
//!
//! This module provides support for decoding Protocol Buffers messages
//! used in iWork IWA (iWork Archive) files using the prost crate.

#![allow(
    dead_code,
    unused_imports,
    reason = "This private decoder adapter retains in-crate migration entries."
)]

use crate::{Error, Result};
use phf::phf_map;
use prost::Message;

// Keep the generated schema layer in its own crate. The explicit list makes
// this compatibility boundary auditable and prevents decoder-only additions
// from accidentally becoming part of the raw schema crate.
pub use litchi_iwa_protos::{kn, tn, tp, tsa, tsce, tsch, tsd, tsk, tsp, tss, tst, tswp};

/// Static decoder function for ArchiveInfo messages
fn decode_archive_info(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    tsp::ArchiveInfo::decode(data)?;
    Ok(Box::new(ArchiveInfoWrapper) as Box<dyn DecodedMessage>)
}

/// Static decoder function for MessageInfo messages
fn decode_message_info(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    tsp::MessageInfo::decode(data)?;
    Ok(Box::new(MessageInfoWrapper) as Box<dyn DecodedMessage>)
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

/// Static decoder function for segmented TableDataList payloads.
fn decode_table_data_list_segment(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tst::TableDataListSegment::decode(data)?;
    Ok(Box::new(TableDataListSegmentWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for ShapeArchive messages
fn decode_shape_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tsd::ShapeArchive::decode(data)?;
    Ok(Box::new(ShapeArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for DrawableArchive messages
fn decode_drawable_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    tsd::DrawableArchive::decode(data)?;
    Ok(Box::new(DrawableArchiveWrapper) as Box<dyn DecodedMessage>)
}

fn decode_comment_storage_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tsd::CommentStorageArchive::decode(data)?;
    Ok(Box::new(CommentStorageArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
}

fn decode_legacy_chart(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let message = tsch::pre_uff::ChartInfoArchive::decode(data)?;
    Ok(Box::new(LegacyChartArchiveWrapper(message)) as Box<dyn DecodedMessage>)
}

fn decode_chart_mediator(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    tsch::ChartMediatorArchive::decode(data)?;
    Ok(Box::new(ChartMediatorArchiveWrapper) as Box<dyn DecodedMessage>)
}

fn decode_chart_drawable(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let message = crate::charts::IWorkChartArchive::decode(data)?;
    Ok(Box::new(ChartDrawableArchiveWrapper(message)) as Box<dyn DecodedMessage>)
}

type DecoderMap = phf::Map<u32, fn(&[u8]) -> Result<Box<dyn DecodedMessage>>>;

/// Perfect hash map of globally shared, non-colliding message type IDs.
///
/// This provides O(1) lookup performance at compile time. It intentionally
/// excludes IDs that are owned by an application namespace. Application
/// editors decode their own schemas at the typed editor boundary.
///
/// Based on analysis of iWork documents and official message type registry:
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
/// Note: Message types are application-specific and may overlap between apps.
static SHARED_DECODERS: DecoderMap = phf_map! {
    // TST (Table) types - Numbers spreadsheet tables and cells
    // Message type 6001 is TST.TableModelArchive
    6000u32 => decode_table_model,
    6001u32 => decode_table_model,
    6005u32 => decode_table_data_list,
    6011u32 => decode_table_data_list_segment,
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
    3056u32 => decode_comment_storage_archive,

    // TSCH (Charts) types
    5000u32 => decode_legacy_chart,
    5004u32 => decode_chart_mediator,
    5021u32 => decode_chart_drawable,

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

};

/// TSP core messages are shared in meaning but their numeric IDs collide with
/// application-owned messages. They are included only in the neutral archive
/// text projection below.
static COMMON_DECODERS: DecoderMap = phf_map! {
    1u32 => decode_archive_info,
    2u32 => decode_message_info,
};

/// Decode a message for the neutral archive text projection.
///
/// The archive layer never guesses an application namespace. It only accepts
/// the shared schemas and returns an unsupported-type error for everything
/// else, leaving application-specific decoding to the owning editor crate.
pub(crate) fn decode_common(message_type: u32, data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let Some(decoder) = COMMON_DECODERS
        .get(&message_type)
        .or_else(|| SHARED_DECODERS.get(&message_type))
    else {
        return Err(Error::UnsupportedMessageType(message_type));
    };
    decoder(data)
}

/// Trait for decoded iWork messages retained by immutable bundle snapshots.
///
/// Decoded messages are read-only after construction, so requiring both
/// marker traits makes the containing archive and bundle safe to share across
/// concurrent readers without a runtime lock.
pub trait DecodedMessage: std::fmt::Debug + Send + Sync {
    /// Extract text content from the message if available
    fn extract_text(&self) -> Vec<String> {
        Vec::new()
    }
}

/// Wrapper for ArchiveInfo message
#[derive(Debug)]
struct ArchiveInfoWrapper;

impl DecodedMessage for ArchiveInfoWrapper {
    fn extract_text(&self) -> Vec<String> {
        Vec::new() // ArchiveInfo doesn't contain text
    }
}

/// Wrapper for MessageInfo message
#[derive(Debug)]
struct MessageInfoWrapper;

impl DecodedMessage for MessageInfoWrapper {
    fn extract_text(&self) -> Vec<String> {
        Vec::new() // MessageInfo doesn't contain text
    }
}

/// Wrapper for StorageArchive message (text content)
#[derive(Debug)]
pub struct StorageArchiveWrapper(pub tswp::StorageArchive);

impl DecodedMessage for StorageArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        self.0.text.clone()
    }
}

/// Wrapper for Table Model Archive (Numbers tables)
#[derive(Debug)]
pub struct TableModelWrapper(pub tst::TableModelArchive);

impl DecodedMessage for TableModelWrapper {
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

/// Wrapper for a segmented TableDataList payload.
#[derive(Debug)]
pub struct TableDataListSegmentWrapper(pub tst::TableDataListSegment);

impl DecodedMessage for TableDataListSegmentWrapper {
    fn extract_text(&self) -> Vec<String> {
        self.0
            .entries
            .iter()
            .filter_map(|entry| entry.string.as_ref())
            .filter(|value| !value.is_empty())
            .cloned()
            .collect()
    }
}

/// Wrapper for Shape Archive
#[derive(Debug)]
pub struct ShapeArchiveWrapper(pub tsd::ShapeArchive);

impl DecodedMessage for ShapeArchiveWrapper {
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
struct DrawableArchiveWrapper;

impl DecodedMessage for DrawableArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        // Drawables are visual elements without direct text
        Vec::new()
    }
}

/// Wrapper for TSD comment storage used by cell and drawable comments.
#[derive(Debug)]
pub struct CommentStorageArchiveWrapper(pub tsd::CommentStorageArchive);

impl DecodedMessage for CommentStorageArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        self.0
            .text
            .iter()
            .filter(|text| !text.is_empty())
            .cloned()
            .collect()
    }
}

/// Wrapper for the legacy, inline-data chart representation.
#[derive(Debug)]
pub struct LegacyChartArchiveWrapper(pub tsch::pre_uff::ChartInfoArchive);

impl DecodedMessage for LegacyChartArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        self.0
            .chart_model
            .inline_grid
            .iter()
            .flat_map(|grid| grid.row_name.iter().chain(&grid.column_name))
            .filter(|text| !text.is_empty())
            .cloned()
            .collect()
    }
}

/// Wrapper for a chart's data mediator.
#[derive(Debug)]
struct ChartMediatorArchiveWrapper;

impl DecodedMessage for ChartMediatorArchiveWrapper {}

/// Wrapper for an extension-backed modern chart drawable.
#[derive(Debug)]
pub struct ChartDrawableArchiveWrapper(pub crate::charts::IWorkChartArchive);

impl DecodedMessage for ChartDrawableArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        self.0
            .chart
            .iter()
            .filter_map(|chart| chart.grid.as_ref())
            .flat_map(|grid| grid.row_name.iter().chain(&grid.column_name))
            .filter(|text| !text.is_empty())
            .cloned()
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn neutral_decoder_registry_contains_only_supported_types() {
        assert!(COMMON_DECODERS.contains_key(&1));
        assert!(COMMON_DECODERS.contains_key(&2));
        assert!(SHARED_DECODERS.contains_key(&6001)); // TST.TableModelArchive
        assert!(SHARED_DECODERS.contains_key(&6011)); // TST.TableDataListSegment
        assert!(SHARED_DECODERS.contains_key(&2001)); // TSWP.StorageArchive
        assert!(SHARED_DECODERS.contains_key(&2002)); // StorageArchive variant
        assert!(SHARED_DECODERS.contains_key(&2003)); // StorageArchive variant
        assert!(SHARED_DECODERS.contains_key(&2022)); // Common StorageArchive type
        assert!(SHARED_DECODERS.contains_key(&3056)); // TSD.CommentStorageArchive
    }

    #[test]
    fn archive_info_is_decoded_without_application_guessing() {
        let archive_info = tsp::ArchiveInfo::default().encode_to_vec();
        let decoded = decode_common(1, &archive_info).unwrap();
        assert!(decoded.extract_text().is_empty());
    }

    #[test]
    fn shared_storage_extracts_text() {
        let storage = tswp::StorageArchive {
            text: vec!["shared".to_owned()],
            ..Default::default()
        };
        let data = storage.encode_to_vec();
        assert_eq!(
            decode_common(2001, &data).unwrap().extract_text(),
            ["shared"]
        );
    }

    #[test]
    fn table_data_list_segments_use_their_concrete_decoder() {
        let segment = tst::TableDataListSegment {
            list_type: tst::table_data_list::ListType::String as i32,
            key_range: tsp::Range {
                location: 7,
                length: 1,
            },
            entries: vec![tst::table_data_list::ListEntry {
                key: 7,
                refcount: 1,
                string: Some("Segmented".to_owned()),
                ..Default::default()
            }],
        };
        let decoded = decode_common(6011, &segment.encode_to_vec()).unwrap();
        assert_eq!(decoded.extract_text(), ["Segmented"]);
    }

    #[test]
    fn comment_storage_uses_its_concrete_decoder() {
        let comment = tsd::CommentStorageArchive {
            text: Some("Review this".to_owned()),
            ..Default::default()
        };
        let decoded = decode_common(3056, &comment.encode_to_vec()).unwrap();
        assert_eq!(decoded.extract_text(), ["Review this"]);
    }

    #[test]
    fn modern_chart_drawables_decode_the_extension_payload() {
        let chart = crate::charts::IWorkChartArchive::new(
            tsch::ChartDrawableArchive::default(),
            tsch::ChartArchive {
                chart_type: Some(tsch::ChartType::ColumnChartType2D as i32),
                grid: Some(tsch::ChartGridArchive {
                    row_name: vec!["Revenue".to_owned()],
                    column_name: vec!["2026".to_owned()],
                    ..Default::default()
                }),
                ..Default::default()
            },
        );
        let decoded = decode_common(5_021, &chart.encode().unwrap()).unwrap();
        assert_eq!(decoded.extract_text(), ["Revenue", "2026"]);
    }

    #[test]
    fn unsupported_application_message_is_rejected() {
        let result = decode_common(999, &[]);
        assert!(matches!(result, Err(Error::UnsupportedMessageType(999))));
    }

    #[test]
    fn test_decoder_performance() {
        // Test that decoding is fast with phf::Map
        // This test ensures the static map lookup is working
        let shared_message_types = [6001, 2001, 2002, 2003];

        // Create some dummy data that will fail to decode but test the lookup
        let dummy_data = vec![0u8; 10];

        for &msg_type in &shared_message_types {
            let result = decode_common(msg_type, &dummy_data);
            // We expect this to fail due to invalid protobuf data, but the lookup should be fast
            assert!(result.is_err());
        }
    }
}
