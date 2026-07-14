//! Typed wrappers around generated protobuf messages.

use super::{kn, tn, tp, tsch, tsd, tsp, tst, tswp};

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
