// Image extraction from Office documents
//
// This module provides high-level functionality to extract all images
// from Microsoft Office documents (PPT, DOC) using the Escher drawing layer.

use crate::common::error::Result;
use crate::images::{Blip, BlipStore, BlipStoreEntry};
use crate::ole::ppt::escher::{EscherContainer, EscherParser, EscherRecordType};
use std::borrow::Cow;

/// Extracted image with metadata
#[derive(Debug, Clone)]
pub struct ExtractedImage<'data> {
    /// The parsed BLIP data
    pub blip: Blip<'data>,
    /// Optional name/filename hint
    pub name: Option<String>,
    /// Index in the document's image collection
    pub index: usize,
}

impl<'data> ExtractedImage<'data> {
    /// Create a new extracted image
    pub fn new(blip: Blip<'data>, name: Option<String>, index: usize) -> Self {
        Self { blip, name, index }
    }

    /// Get the BLIP type
    pub fn blip_type(&self) -> Option<crate::images::BlipType> {
        self.blip.blip_type()
    }

    /// Get file extension for this image
    pub fn extension(&self) -> &'static str {
        self.blip_type().map(|t| t.extension()).unwrap_or("bin")
    }

    /// Get suggested filename
    pub fn suggested_filename(&self) -> String {
        if let Some(name) = &self.name {
            // Check if name already has extension
            if name.contains('.') {
                name.clone()
            } else {
                format!("{}.{}", name, self.extension())
            }
        } else {
            format!("image_{:03}.{}", self.index, self.extension())
        }
    }

    /// Get raw picture data
    pub fn raw_data(&self) -> &[u8] {
        self.blip.picture_data()
    }

    /// Get decompressed picture data
    pub fn decompressed_data(&self) -> Result<Cow<'data, [u8]>> {
        self.blip.get_decompressed_data()
    }

    /// Convert to PNG format
    #[cfg(feature = "imgconv")]
    pub fn to_png(&self, width: Option<u32>, height: Option<u32>) -> Result<Vec<u8>> {
        crate::images::convert_blip_to_png(&self.blip, width, height)
    }

    /// Convert to JPEG format
    #[cfg(feature = "imgconv")]
    pub fn to_jpeg(&self, width: Option<u32>, height: Option<u32>) -> Result<Vec<u8>> {
        crate::images::convert_blip_to_jpeg(&self.blip, width, height)
    }
}

/// Image extractor for Office documents
///
/// This provides functionality to extract all images from Office documents
/// by parsing the Escher drawing layer and BLIP records.
pub struct ImageExtractor;

impl ImageExtractor {
    /// Extract BLIP store (BSE index) from Escher drawing data
    ///
    /// # Arguments
    /// * `data` - Escher drawing data (typically from Drawing Group Container)
    ///
    /// # Returns
    /// BlipStore containing all BSE entries
    pub fn extract_blip_store<'data>(data: &'data [u8]) -> Result<BlipStore<'data>> {
        let mut store = BlipStore::new();

        // Parse Escher records
        let parser = EscherParser::new(data);
        for record in parser.records() {
            let record = record?;

            // Look for BStoreContainer (0xF001)
            if record.record_type == EscherRecordType::BStoreContainer {
                // Parse container to get BSE records
                let container = EscherContainer::new(record);

                // The instance field in BStoreContainer header indicates the number of BLIPs
                let blip_count = container.record().instance as usize;
                store = BlipStore::with_capacity(blip_count);

                // Each child record should be a BSE (0xF007)
                for child in container.children().flatten() {
                    if child.record_type == EscherRecordType::BSE {
                        match BlipStoreEntry::parse(child.data) {
                            Ok(bse) => store.add_entry(bse),
                            Err(e) => {
                                // Log error but continue processing
                                eprintln!("Warning: Failed to parse BSE entry: {}", e);
                            },
                        }
                    }
                }

                break; // Found the store, no need to continue
            }
        }

        Ok(store)
    }

    /// Extract all BLIPs from Escher drawing data
    ///
    /// This extracts actual BLIP records (image data) from the drawing layer.
    ///
    /// # Arguments
    /// * `data` - Escher drawing data
    ///
    /// # Returns
    /// Vector of extracted images with metadata
    ///
    /// Note: Returns owned BLIPs (static lifetime) since we reconstruct data from records
    pub fn extract_blips(data: &[u8]) -> Result<Vec<ExtractedImage<'static>>> {
        let mut images = Vec::new();
        let mut index = 0;

        // First, try to extract the BLIP store for metadata
        let store = Self::extract_blip_store(data).ok();

        // Parse all Escher records looking for BLIPs
        let parser = EscherParser::new(data);
        for record in parser.records() {
            let record = record?;

            // Check if this is a BLIP record
            let is_blip = matches!(
                record.record_type,
                EscherRecordType::BlipEmf
                    | EscherRecordType::BlipWmf
                    | EscherRecordType::BlipPict
                    | EscherRecordType::BlipJpeg
                    | EscherRecordType::BlipPng
                    | EscherRecordType::BlipDib
                    | EscherRecordType::BlipTiff
            );

            if is_blip {
                // Need to reconstruct full BLIP record with header
                let mut full_data = Vec::with_capacity(8 + record.data.len());

                // Reconstruct header
                let ver_inst = (record.instance << 4) | (record.version as u16);
                full_data.extend_from_slice(&ver_inst.to_le_bytes());
                full_data.extend_from_slice(&record.record_type_raw.to_le_bytes());
                full_data.extend_from_slice(&record.length.to_le_bytes());
                full_data.extend_from_slice(record.data);

                // Parse the BLIP and immediately convert to owned
                match Blip::parse(&full_data) {
                    Ok(blip) => {
                        // Convert to owned to avoid lifetime issues
                        let owned_blip = blip.into_owned();

                        // Try to get name from store if available
                        let name = if let Some(ref blip_store) = store {
                            blip_store
                                .get_entry(index)
                                .and_then(|bse| bse.name.as_ref().map(|n| n.to_string()))
                        } else {
                            None
                        };

                        images.push(ExtractedImage::new(owned_blip, name, index));
                        index += 1;
                    },
                    Err(e) => {
                        eprintln!("Warning: Failed to parse BLIP at index {}: {}", index, e);
                    },
                }
            }
        }

        Ok(images)
    }

    /// Extract images from a specific Escher container
    ///
    /// This is useful when you want to extract images from a specific
    /// part of a document (e.g., a specific slide in PPT).
    pub fn extract_from_container(
        container: &EscherContainer,
    ) -> Result<Vec<ExtractedImage<'static>>> {
        let mut images = Vec::new();
        let mut index = 0;

        // Recursively search for BLIP records
        Self::extract_from_container_recursive(container, &mut images, &mut index)?;

        Ok(images)
    }

    /// Recursively extract BLIPs from a container and its children
    fn extract_from_container_recursive(
        container: &EscherContainer,
        images: &mut Vec<ExtractedImage<'static>>,
        index: &mut usize,
    ) -> Result<()> {
        // Check if this container has BLIP records
        for child_result in container.children() {
            let child = match child_result {
                Ok(c) => c,
                Err(_) => continue, // Skip invalid records
            };
            let is_blip = matches!(
                child.record_type,
                EscherRecordType::BlipEmf
                    | EscherRecordType::BlipWmf
                    | EscherRecordType::BlipPict
                    | EscherRecordType::BlipJpeg
                    | EscherRecordType::BlipPng
                    | EscherRecordType::BlipDib
                    | EscherRecordType::BlipTiff
            );

            if is_blip {
                // Reconstruct full BLIP record
                let mut full_data = Vec::with_capacity(8 + child.data.len());
                let ver_inst = (child.instance << 4) | (child.version as u16);
                full_data.extend_from_slice(&ver_inst.to_le_bytes());
                full_data.extend_from_slice(&child.record_type_raw.to_le_bytes());
                full_data.extend_from_slice(&child.length.to_le_bytes());
                full_data.extend_from_slice(child.data);

                if let Ok(blip) = Blip::parse(&full_data) {
                    // Convert to owned to avoid lifetime issues
                    let owned_blip = blip.into_owned();
                    images.push(ExtractedImage::new(owned_blip, None, *index));
                    *index += 1;
                }
            } else if child.is_container() {
                // Recurse into child containers
                let child_container = EscherContainer::new(child);
                Self::extract_from_container_recursive(&child_container, images, index)?;
            }
        }

        Ok(())
    }

    /// Extract images from Pictures stream (PPT specific)
    ///
    /// In PPT files, images are often stored in a separate "Pictures" stream.
    /// This method extracts all BLIPs from that stream.
    ///
    /// # Arguments
    /// * `pictures_data` - Raw data from the Pictures stream
    ///
    /// # Returns
    /// Vector of extracted images
    pub fn extract_from_pictures_stream(
        pictures_data: &[u8],
    ) -> Result<Vec<ExtractedImage<'static>>> {
        Self::extract_blips(pictures_data)
    }
}

/// High-level image extraction from PPT presentations
#[cfg(feature = "ole")]
pub mod ppt {
    use super::*;
    use crate::ole::OleFile;
    use std::io::{Read, Seek};

    impl ImageExtractor {
        /// Extract all images from a PPT presentation
        ///
        /// # Arguments
        /// * `ole` - Opened OLE file for the PPT presentation
        ///
        /// # Returns
        /// Vector of all extracted images
        pub fn extract_from_ppt<R: Read + Seek>(
            ole: &mut OleFile<R>,
        ) -> Result<Vec<ExtractedImage<'static>>> {
            let mut all_images = Vec::new();

            // Try to read Pictures stream
            if ole.exists(&["Pictures"]) {
                match ole.open_stream(&["Pictures"]) {
                    Ok(data) => {
                        let images = Self::extract_from_pictures_stream(&data)?;
                        // Convert to owned data since we're returning from function
                        all_images.extend(images.into_iter().map(|img| ExtractedImage {
                            blip: img.blip.into_owned(),
                            name: img.name,
                            index: img.index,
                        }));
                    },
                    Err(e) => {
                        eprintln!("Warning: Failed to read Pictures stream: {}", e);
                    },
                }
            }

            // Also check PowerPoint Document stream for embedded drawings
            if ole.exists(&["PowerPoint Document"]) {
                match ole.open_stream(&["PowerPoint Document"]) {
                    Ok(data) => {
                        let images = Self::extract_blips(&data)?;
                        let offset = all_images.len();
                        all_images.extend(images.into_iter().map(|mut img| {
                            img.index += offset;
                            ExtractedImage {
                                blip: img.blip.into_owned(),
                                name: img.name,
                                index: img.index,
                            }
                        }));
                    },
                    Err(e) => {
                        eprintln!("Warning: Failed to read PowerPoint Document stream: {}", e);
                    },
                }
            }

            Ok(all_images)
        }
    }
}

/// High-level image extraction from DOC documents
#[cfg(feature = "ole")]
pub mod doc {
    use super::*;
    use crate::ole::OleFile;
    use std::io::{Read, Seek};

    impl ImageExtractor {
        /// Extract all images from a DOC document
        ///
        /// # Arguments
        /// * `ole` - Opened OLE file for the DOC document
        ///
        /// # Returns
        /// Vector of all extracted images
        pub fn extract_from_doc<R: Read + Seek>(
            ole: &mut OleFile<R>,
        ) -> Result<Vec<ExtractedImage<'static>>> {
            let mut all_images = Vec::new();

            // Try to read Data stream (contains embedded objects and images)
            if ole.exists(&["Data"]) {
                match ole.open_stream(&["Data"]) {
                    Ok(data) => {
                        let images = Self::extract_blips(&data)?;
                        all_images.extend(images.into_iter().map(|img| ExtractedImage {
                            blip: img.blip.into_owned(),
                            name: img.name,
                            index: img.index,
                        }));
                    },
                    Err(e) => {
                        eprintln!("Warning: Failed to read Data stream: {}", e);
                    },
                }
            }

            // Also check 1Table/0Table streams for drawing data
            for table_name in &["1Table", "0Table"] {
                if ole.exists(&[table_name]) {
                    match ole.open_stream(&[table_name]) {
                        Ok(data) => {
                            let images = Self::extract_blips(&data)?;
                            let offset = all_images.len();
                            all_images.extend(images.into_iter().map(|mut img| {
                                img.index += offset;
                                ExtractedImage {
                                    blip: img.blip.into_owned(),
                                    name: img.name,
                                    index: img.index,
                                }
                            }));
                        },
                        Err(e) => {
                            eprintln!("Warning: Failed to read {} stream: {}", table_name, e);
                        },
                    }
                }
            }

            Ok(all_images)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_extracted_image_filename() {
        let blip_data = vec![
            0x0A, 0x00, // version=0, instance=0
            0x1D, 0xF0, // JPEG BLIP
            0x19, 0x00, 0x00, 0x00, // length = 25
            // UID (16 bytes)
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E,
            0x0F, 0x10, 0xFF, // marker
            // Minimal JPEG data
            0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46,
        ];

        let blip = Blip::parse(&blip_data).unwrap();
        let img = ExtractedImage::new(blip, None, 0);

        assert_eq!(img.extension(), "jpg");
        assert_eq!(img.suggested_filename(), "image_000.jpg");

        let img_with_name = ExtractedImage::new(
            Blip::parse(&blip_data).unwrap(),
            Some("photo".to_string()),
            5,
        );
        assert_eq!(img_with_name.suggested_filename(), "photo.jpg");
    }
}
