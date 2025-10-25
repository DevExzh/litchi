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

    /// Get the recommended output extension for this image
    ///
    /// - EMF, WMF: Converted to SVG
    /// - PICT: Converted to PNG (SVG not yet implemented)
    /// - Bitmaps (PNG, JPEG, etc.): Keep their original extension
    pub fn output_extension(&self) -> &'static str {
        use crate::images::BlipType;

        if let Some(blip_type) = self.blip_type() {
            match blip_type {
                BlipType::Emf | BlipType::Wmf => "svg",
                BlipType::Pict => "png", // PICT to SVG not yet implemented
                _ => blip_type.extension(),
            }
        } else {
            "bin"
        }
    }

    /// Get suggested filename for output
    ///
    /// Uses the output extension (SVG for metafiles, original extension for bitmaps)
    pub fn suggested_filename(&self) -> String {
        if let Some(name) = &self.name {
            // Check if name already has extension
            if name.contains('.') {
                // Replace extension with output extension
                let stem = name.split('.').next().unwrap_or(name);
                format!("{}.{}", stem, self.output_extension())
            } else {
                format!("{}.{}", name, self.output_extension())
            }
        } else {
            format!("image_{:03}.{}", self.index, self.output_extension())
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

    /// Convert metafile to SVG format
    ///
    /// For EMF and WMF formats, converts to SVG.
    /// PICT format is not yet supported for SVG conversion and will return an error.
    /// Returns an error if the image is not a metafile.
    #[cfg(feature = "imgconv")]
    pub fn to_svg(&self) -> Result<String> {
        use crate::images::BlipType;

        // For WMF, we need to add the placeable header using BLIP metadata
        // For EMF and others, just use decompressed data
        let data = self.blip.get_picture_data_for_conversion()?;

        match self.blip_type() {
            Some(BlipType::Emf) => crate::images::emf::convert_emf_to_svg(&data),
            Some(BlipType::Wmf) => crate::images::wmf::convert_wmf_to_svg(&data),
            Some(BlipType::Pict) => {
                // PICT to SVG conversion is not yet implemented
                // Fall back to error for now
                Err(crate::common::error::Error::ParseError(
                    "PICT to SVG conversion is not yet implemented".into(),
                ))
            },
            _ => Err(crate::common::error::Error::ParseError(
                "Image is not a metafile format (EMF/WMF/PICT)".into(),
            )),
        }
    }

    /// Extract the image in its recommended format
    ///
    /// - For metafiles (EMF, WMF): Converts to SVG
    /// - For PICT: Converts to PNG (SVG not yet implemented)
    /// - For bitmaps (PNG, JPEG, DIB, TIFF): Returns raw decompressed data
    ///
    /// This is the recommended method for extracting images as it preserves
    /// vector graphics as vectors and bitmaps as bitmaps.
    #[cfg(feature = "imgconv")]
    pub fn extract(&self) -> Result<Vec<u8>> {
        use crate::images::BlipType;

        if let Some(blip_type) = self.blip_type() {
            match blip_type {
                BlipType::Emf | BlipType::Wmf => {
                    // Convert EMF/WMF to SVG
                    self.to_svg().map(|s| s.into_bytes())
                },
                BlipType::Pict => {
                    // PICT to SVG not yet implemented, convert to PNG instead
                    self.to_png(Some(800), None)
                },
                _ => {
                    // Extract bitmaps as-is
                    self.decompressed_data().map(|cow| cow.into_owned())
                },
            }
        } else {
            // Unknown format, just return decompressed data
            self.decompressed_data().map(|cow| cow.into_owned())
        }
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

    /// Search for BLIP records in raw data (for DOC files)
    ///
    /// In DOC files, the Data stream may contain BLIP records at various offsets,
    /// not necessarily starting at the beginning. This function searches for
    /// BLIP record signatures throughout the data.
    ///
    /// # Arguments
    /// * `data` - Raw data to search (typically from the Data stream)
    ///
    /// # Returns
    /// Vector of extracted images
    fn search_blips_in_data(data: &[u8]) -> Result<Vec<ExtractedImage<'static>>> {
        let mut images = Vec::new();
        let mut index = 0;

        // BLIP record type IDs to search for
        const BLIP_SIGNATURES: &[(u16, &str)] = &[
            (0xF01A, "emf"),
            (0xF01B, "wmf"),
            (0xF01C, "pict"),
            (0xF01D, "jpeg"),
            (0xF01E, "png"),
            (0xF01F, "dib"),
            (0xF029, "tiff"),
        ];

        // Search through the data for BLIP signatures
        let mut offset = 0;
        while offset + 8 <= data.len() {
            // Read potential record header
            if offset + 4 <= data.len() {
                let record_type = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);

                // Check if this looks like a BLIP record
                let is_blip = BLIP_SIGNATURES.iter().any(|(sig, _)| *sig == record_type);

                if is_blip && offset + 8 <= data.len() {
                    // Read the length
                    let length = u32::from_le_bytes([
                        data[offset + 4],
                        data[offset + 5],
                        data[offset + 6],
                        data[offset + 7],
                    ]) as usize;

                    // Validate length is reasonable
                    if length > 0 && length < 100_000_000 && offset + 8 + length <= data.len() {
                        // Extract the full BLIP record
                        let blip_data = &data[offset..offset + 8 + length];

                        // Try to parse it
                        match Blip::parse(blip_data) {
                            Ok(blip) => {
                                images.push(ExtractedImage::new(blip.into_owned(), None, index));
                                index += 1;
                                // Skip past this record
                                offset += 8 + length;
                                continue;
                            },
                            Err(_) => {
                                // Not a valid BLIP, continue searching
                            },
                        }
                    }
                }
            }

            offset += 1;
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
        /// In DOC files, images are typically stored in the Data stream as raw BLIP records.
        /// The table stream contains metadata about where these images are referenced in the text,
        /// but the actual image data is in the Data stream or embedded in the ObjectPool.
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
            // In DOC files, this is where the actual picture data is stored
            if ole.exists(&["Data"]) {
                match ole.open_stream(&["Data"]) {
                    Ok(data) => {
                        // The Data stream may contain multiple BLIP records at various offsets
                        // Use the search function to find them all
                        match Self::search_blips_in_data(&data) {
                            Ok(images) => {
                                all_images.extend(images);
                            },
                            Err(e) => {
                                eprintln!(
                                    "Warning: Failed to search for BLIPs in Data stream: {}",
                                    e
                                );
                            },
                        }
                    },
                    Err(e) => {
                        eprintln!("Warning: Failed to read Data stream: {}", e);
                    },
                }
            }

            // Note: We don't try to parse the entire table stream as Escher data
            // because it contains various other structures and the drawing data
            // is at specific offsets that would need to be parsed from the FIB.
            // For most practical purposes, the Data stream contains the images we need.

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
