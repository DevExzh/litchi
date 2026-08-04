//! Protobuf Message Support for iWork IWA Files
//!
//! This module provides support for decoding Protocol Buffers messages
//! used in iWork IWA (iWork Archive) files using the prost crate.

use crate::{Error, Result};
use phf::phf_map;
use prost::Message;

// Keep the generated schema layer in its own crate. The explicit list makes
// this compatibility boundary auditable and prevents decoder-only additions
// from accidentally becoming part of the raw schema crate.
pub use litchi_iwa_protos::{
    kn, knsos, tn, tnsos, tp, tpsos, tsa, tsasos, tsce, tsch, tschsos, tsck, tscksos, tsd, tsdsos,
    tsk, tsp, tss, tsssos, tst, tstsos, tswp, tswpsos,
};

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

fn decode_pages_theme(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tp::ThemeArchive::decode(data)?;
    Ok(Box::new(PagesThemeWrapper(msg)) as Box<dyn DecodedMessage>)
}

fn decode_pages_section(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tp::SectionArchive::decode(data)?;
    Ok(Box::new(PagesSectionWrapper(msg)) as Box<dyn DecodedMessage>)
}

fn decode_pages_section_template(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tp::SectionTemplateArchive::decode(data)?;
    Ok(Box::new(PagesSectionTemplateWrapper(msg)) as Box<dyn DecodedMessage>)
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

fn decode_keynote_build(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = kn::BuildArchive::decode(data)?;
    Ok(Box::new(KeynoteBuildWrapper(msg)) as Box<dyn DecodedMessage>)
}

fn decode_keynote_build_chunk(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = kn::BuildChunkArchive::decode(data)?;
    Ok(Box::new(KeynoteBuildChunkWrapper(msg)) as Box<dyn DecodedMessage>)
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
    let msg = tsd::DrawableArchive::decode(data)?;
    Ok(Box::new(DrawableArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
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
    let message = tsch::ChartMediatorArchive::decode(data)?;
    Ok(Box::new(ChartMediatorArchiveWrapper(message)) as Box<dyn DecodedMessage>)
}

fn decode_chart_drawable(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let message = crate::charts::IWorkChartArchive::decode(data)?;
    Ok(Box::new(ChartDrawableArchiveWrapper(message)) as Box<dyn DecodedMessage>)
}

type DecoderMap = phf::Map<u32, fn(&[u8]) -> Result<Box<dyn DecodedMessage>>>;

/// Typed application namespace used to resolve colliding iWork message IDs.
///
/// This is a renamed view of the crate's canonical application enum. Keeping
/// the context as an enum makes it impossible for callers to accidentally pass
/// an unvalidated string or an arbitrary numeric application identifier to the
/// decoder boundary.
pub use crate::registry::Application as ApplicationDecodeContext;

/// Perfect hash map of globally shared, non-colliding message type IDs.
///
/// This provides O(1) lookup performance at compile time. It intentionally
/// excludes IDs that are also owned by an application namespace, even when a
/// common schema is available for that ID. Those IDs must go through
/// [`decode_with_context`].
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
/// application-owned messages. They therefore remain available only through
/// an explicit [`ApplicationDecodeContext::Common`] context.
static COMMON_DECODERS: DecoderMap = phf_map! {
    1u32 => decode_archive_info,
    2u32 => decode_message_info,
};

/// Pages-specific decoder table. The selected application is part of the
/// dispatch key; no other application table is consulted on a miss.
static PAGES_DECODERS: DecoderMap = phf_map! {
    10000u32 => decode_pages_document,
    10001u32 => decode_pages_theme,
    10011u32 => decode_pages_section,
    10143u32 => decode_pages_section_template,
};

/// Numbers-specific decoder table. Type 3 is intentionally not present in the
/// context-free map because low TN IDs overlap with other iWork namespaces.
static NUMBERS_DECODERS: DecoderMap = phf_map! {
    3u32 => decode_numbers_sheet,       // TN.FormBasedSheetArchive
};

/// Keynote-specific decoder table. Low KN IDs are never selected without an
/// explicit Keynote context.
static KEYNOTE_DECODERS: DecoderMap = phf_map! {
    5u32 => decode_keynote_slide,       // KN.SlideArchive
    6u32 => decode_keynote_slide,       // KN.SlideArchive (variant)
    8u32 => decode_keynote_build,       // KN.BuildArchive
    153u32 => decode_keynote_build_chunk, // KN.BuildChunkArchive
};

fn decoder_for_context(
    context: ApplicationDecodeContext,
    message_type: u32,
) -> Option<fn(&[u8]) -> Result<Box<dyn DecodedMessage>>> {
    let decoder = match context {
        // Common is the only namespace allowed to opt into the shared table.
        // Application contexts remain exact maps so a wrong-context ID cannot
        // succeed merely because its wire bytes fit a shared schema.
        ApplicationDecodeContext::Common => COMMON_DECODERS
            .get(&message_type)
            .or_else(|| SHARED_DECODERS.get(&message_type)),
        ApplicationDecodeContext::Pages => PAGES_DECODERS.get(&message_type),
        ApplicationDecodeContext::Numbers => NUMBERS_DECODERS.get(&message_type),
        ApplicationDecodeContext::Keynote => KEYNOTE_DECODERS.get(&message_type),
    };

    decoder.copied()
}

/// Decode a globally shared message without an application context.
///
/// Ambiguous IDs and application-owned IDs are rejected here. Callers that
/// know the owning application must use [`decode_with_context`] instead.
pub fn decode(message_type: u32, data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    if let Some(decoder) = SHARED_DECODERS.get(&message_type) {
        decoder(data)
    } else {
        Err(Error::UnsupportedMessageType(message_type))
    }
}

/// Decode a message using an explicit typed application namespace.
///
/// Dispatch is performed by one compile-time perfect-hash table for the
/// selected application. Only the explicit `Common` namespace may select the
/// shared table. A miss is terminal: the decoder never tries another
/// application, inspects payload bytes to guess a schema, or falls back to a
/// permissive protobuf parse.
pub fn decode_with_context(
    context: ApplicationDecodeContext,
    message_type: u32,
    data: &[u8],
) -> Result<Box<dyn DecodedMessage>> {
    let Some(decoder) = decoder_for_context(context, message_type) else {
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
        10000
    }

    fn extract_text(&self) -> Vec<String> {
        Vec::new() // Document metadata doesn't contain direct text
    }
}

/// Wrapper for TP.ThemeArchive.
#[derive(Debug)]
pub struct PagesThemeWrapper(pub tp::ThemeArchive);

impl DecodedMessage for PagesThemeWrapper {
    fn message_type(&self) -> u32 {
        10001
    }
}

/// Wrapper for TP.SectionArchive.
#[derive(Debug)]
pub struct PagesSectionWrapper(pub tp::SectionArchive);

impl DecodedMessage for PagesSectionWrapper {
    fn message_type(&self) -> u32 {
        10011
    }

    fn extract_text(&self) -> Vec<String> {
        self.0
            .name
            .iter()
            .filter(|name| !name.is_empty())
            .cloned()
            .collect()
    }
}

/// Wrapper for TP.SectionTemplateArchive.
#[derive(Debug)]
pub struct PagesSectionTemplateWrapper(pub tp::SectionTemplateArchive);

impl DecodedMessage for PagesSectionTemplateWrapper {
    fn message_type(&self) -> u32 {
        10143
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

/// Wrapper for a Keynote object build.
#[derive(Debug)]
pub struct KeynoteBuildWrapper(pub kn::BuildArchive);

impl DecodedMessage for KeynoteBuildWrapper {
    fn message_type(&self) -> u32 {
        8
    }
}

/// Wrapper for a Keynote build timing chunk.
#[derive(Debug)]
pub struct KeynoteBuildChunkWrapper(pub kn::BuildChunkArchive);

impl DecodedMessage for KeynoteBuildChunkWrapper {
    fn message_type(&self) -> u32 {
        153
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

/// Wrapper for a segmented TableDataList payload.
#[derive(Debug)]
pub struct TableDataListSegmentWrapper(pub tst::TableDataListSegment);

impl DecodedMessage for TableDataListSegmentWrapper {
    fn message_type(&self) -> u32 {
        6011
    }

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

/// Wrapper for TSD comment storage used by cell and drawable comments.
#[derive(Debug)]
pub struct CommentStorageArchiveWrapper(pub tsd::CommentStorageArchive);

impl DecodedMessage for CommentStorageArchiveWrapper {
    fn message_type(&self) -> u32 {
        3056
    }

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
    fn message_type(&self) -> u32 {
        5_000
    }

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
pub struct ChartMediatorArchiveWrapper(pub tsch::ChartMediatorArchive);

impl DecodedMessage for ChartMediatorArchiveWrapper {
    fn message_type(&self) -> u32 {
        5_004
    }
}

/// Wrapper for an extension-backed modern chart drawable.
#[derive(Debug)]
pub struct ChartDrawableArchiveWrapper(pub crate::charts::IWorkChartArchive);

impl DecodedMessage for ChartDrawableArchiveWrapper {
    fn message_type(&self) -> u32 {
        5_021
    }

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
    fn test_message_decoder_creation() {
        // Shared dispatch must not expose colliding application IDs.
        assert!(!SHARED_DECODERS.contains_key(&1));
        assert!(!SHARED_DECODERS.contains_key(&2));
        assert!(!SHARED_DECODERS.contains_key(&3));
        assert!(!SHARED_DECODERS.contains_key(&5));
        assert!(COMMON_DECODERS.contains_key(&1));
        assert!(COMMON_DECODERS.contains_key(&2));
        assert!(SHARED_DECODERS.contains_key(&6001)); // TST.TableModelArchive
        assert!(SHARED_DECODERS.contains_key(&6011)); // TST.TableDataListSegment
        assert!(SHARED_DECODERS.contains_key(&2001)); // TSWP.StorageArchive
        assert!(SHARED_DECODERS.contains_key(&2002)); // StorageArchive variant
        assert!(SHARED_DECODERS.contains_key(&2003)); // StorageArchive variant
        assert!(SHARED_DECODERS.contains_key(&2022)); // Common StorageArchive type
        assert!(SHARED_DECODERS.contains_key(&3056)); // TSD.CommentStorageArchive
        assert!(PAGES_DECODERS.contains_key(&10000)); // TP.DocumentArchive
        assert!(PAGES_DECODERS.contains_key(&10001)); // TP.ThemeArchive
        assert!(PAGES_DECODERS.contains_key(&10011)); // TP.SectionArchive
        assert!(PAGES_DECODERS.contains_key(&10143)); // TP.SectionTemplateArchive
        assert!(NUMBERS_DECODERS.contains_key(&3)); // TN.FormBasedSheetArchive
        assert!(KEYNOTE_DECODERS.contains_key(&5)); // KN.SlideArchive
        assert!(KEYNOTE_DECODERS.contains_key(&6)); // KN.SlideArchive variant
        assert!(KEYNOTE_DECODERS.contains_key(&8)); // KN.BuildArchive
        assert!(KEYNOTE_DECODERS.contains_key(&153)); // KN.BuildChunkArchive
    }

    #[test]
    fn ambiguous_ids_require_an_explicit_context() {
        let archive_info = tsp::ArchiveInfo::default().encode_to_vec();

        assert!(matches!(
            decode(1, &archive_info),
            Err(Error::UnsupportedMessageType(1))
        ));
        for context in [
            ApplicationDecodeContext::Pages,
            ApplicationDecodeContext::Numbers,
            ApplicationDecodeContext::Keynote,
        ] {
            assert!(matches!(
                decode_with_context(context, 1, &archive_info),
                Err(Error::UnsupportedMessageType(1))
            ));
        }

        let decoded =
            decode_with_context(ApplicationDecodeContext::Common, 1, &archive_info).unwrap();
        assert_eq!(decoded.message_type(), 1);
    }

    #[test]
    fn application_context_never_falls_back_to_another_namespace() {
        let build = kn::BuildArchive {
            delivery: "All at Once".to_owned(),
            attributes: kn::BuildAttributesArchive::default(),
            ..Default::default()
        };
        let data = build.encode_to_vec();

        assert!(matches!(
            decode_with_context(ApplicationDecodeContext::Pages, 8, &data),
            Err(Error::UnsupportedMessageType(8))
        ));
        assert!(matches!(
            decode_with_context(ApplicationDecodeContext::Numbers, 8, &data),
            Err(Error::UnsupportedMessageType(8))
        ));
        assert_eq!(
            decode_with_context(ApplicationDecodeContext::Keynote, 8, &data)
                .unwrap()
                .message_type(),
            8
        );
    }

    #[test]
    fn shared_ids_are_explicitly_scoped_to_common() {
        let storage = tswp::StorageArchive {
            text: vec!["shared".to_owned()],
            ..Default::default()
        };
        let data = storage.encode_to_vec();

        assert_eq!(
            decode_with_context(ApplicationDecodeContext::Common, 2001, &data)
                .unwrap()
                .extract_text(),
            ["shared"]
        );
        assert!(matches!(
            decode_with_context(ApplicationDecodeContext::Pages, 2001, &data),
            Err(Error::UnsupportedMessageType(2001))
        ));
    }

    #[test]
    fn keynote_build_types_use_their_concrete_decoders() {
        let build = kn::BuildArchive {
            delivery: "All at Once".to_owned(),
            attributes: kn::BuildAttributesArchive::default(),
            ..Default::default()
        };
        assert_eq!(
            decode_with_context(ApplicationDecodeContext::Keynote, 8, &build.encode_to_vec())
                .unwrap()
                .message_type(),
            8
        );

        let chunk = kn::BuildChunkArchive::default();
        assert_eq!(
            decode_with_context(
                ApplicationDecodeContext::Keynote,
                153,
                &chunk.encode_to_vec(),
            )
            .unwrap()
            .message_type(),
            153
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
        let decoded = decode(6011, &segment.encode_to_vec()).unwrap();
        assert_eq!(decoded.message_type(), 6011);
        assert_eq!(decoded.extract_text(), ["Segmented"]);
    }

    #[test]
    fn comment_storage_uses_its_concrete_decoder() {
        let comment = tsd::CommentStorageArchive {
            text: Some("Review this".to_owned()),
            ..Default::default()
        };
        let decoded = decode(3056, &comment.encode_to_vec()).unwrap();
        assert_eq!(decoded.message_type(), 3056);
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
        let decoded = decode(5_021, &chart.encode().unwrap()).unwrap();

        assert_eq!(decoded.message_type(), 5_021);
        assert_eq!(decoded.extract_text(), ["Revenue", "2026"]);
    }

    #[test]
    fn pages_section_types_use_their_concrete_decoders() {
        let section = tp::SectionArchive {
            name: Some("Chapter".to_owned()),
            ..Default::default()
        };
        let decoded = decode_with_context(
            ApplicationDecodeContext::Pages,
            10011,
            &section.encode_to_vec(),
        )
        .unwrap();
        assert_eq!(decoded.message_type(), 10011);
        assert_eq!(decoded.extract_text(), ["Chapter"]);

        let template = tp::SectionTemplateArchive::default();
        let decoded = decode_with_context(
            ApplicationDecodeContext::Pages,
            10143,
            &template.encode_to_vec(),
        )
        .unwrap();
        assert_eq!(decoded.message_type(), 10143);
        assert!(decoded.extract_text().is_empty());
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
        let shared_message_types = [6001, 2001, 2002, 2003];

        // Create some dummy data that will fail to decode but test the lookup
        let dummy_data = vec![0u8; 10];

        for &msg_type in &shared_message_types {
            let result = decode(msg_type, &dummy_data);
            // We expect this to fail due to invalid protobuf data, but the lookup should be fast
            assert!(result.is_err());
        }

        for (context, message_type) in [
            (ApplicationDecodeContext::Pages, 10000),
            (ApplicationDecodeContext::Numbers, 3),
            (ApplicationDecodeContext::Keynote, 5),
        ] {
            let result = decode_with_context(context, message_type, &dummy_data);
            assert!(result.is_err());
        }
    }
}
