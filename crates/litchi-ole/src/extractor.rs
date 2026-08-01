// Image extraction from Office documents
//
// This module provides high-level functionality to extract all images
// from Microsoft Office documents (PPT, DOC) using the Escher drawing layer.

use litchi_core::error::Result;
use litchi_imgconv::{Blip, BlipStore, BlipStoreEntry};
use litchi_odraw::{Container, Parser, Record, RecordKind};
use std::borrow::Cow;

fn odraw_error(error: litchi_odraw::Error) -> litchi_core::error::Error {
    litchi_core::error::Error::ParseError(error.to_string())
}

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
    pub fn blip_type(&self) -> Option<litchi_imgconv::BlipType> {
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
        use litchi_imgconv::BlipType;

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
    pub fn to_png(&self, width: Option<u32>, height: Option<u32>) -> Result<Vec<u8>> {
        litchi_imgconv::convert_blip_to_png(&self.blip, width, height)
    }

    /// Convert to JPEG format
    pub fn to_jpeg(&self, width: Option<u32>, height: Option<u32>) -> Result<Vec<u8>> {
        litchi_imgconv::convert_blip_to_jpeg(&self.blip, width, height)
    }

    /// Convert metafile to SVG format
    ///
    /// For EMF and WMF formats, converts to SVG.
    /// PICT format is not yet supported for SVG conversion and will return an error.
    /// Returns an error if the image is not a metafile.
    pub fn to_svg(&self) -> Result<String> {
        use litchi_imgconv::BlipType;

        // For WMF, we need to add the placeable header using BLIP metadata
        // For EMF and others, just use decompressed data
        let data = self.blip.get_picture_data_for_conversion()?;

        match self.blip_type() {
            Some(BlipType::Emf) => litchi_imgconv::emf::convert_emf_to_svg(&data),
            Some(BlipType::Wmf) => litchi_imgconv::wmf::convert_wmf_to_svg(&data),
            Some(BlipType::Pict) => {
                // PICT to SVG conversion is not yet implemented
                // Fall back to error for now
                Err(litchi_core::error::Error::ParseError(
                    "PICT to SVG conversion is not yet implemented".into(),
                ))
            },
            _ => Err(litchi_core::error::Error::ParseError(
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
    pub fn extract(&self) -> Result<Vec<u8>> {
        use litchi_imgconv::BlipType;

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
        let parser = Parser::new(data);
        for record in parser.records() {
            let record = record.map_err(odraw_error)?;

            // Look for BStoreContainer (0xF001)
            if record.kind() == RecordKind::BStoreContainer {
                // Parse container to get BSE records
                let container = Container::try_new(record).map_err(odraw_error)?;

                // The instance field in BStoreContainer header indicates the number of BLIPs
                let blip_count = usize::from(container.record().instance());
                store = BlipStore::with_capacity(blip_count);

                // Each child record should be a BSE (0xF007)
                for child in container.children() {
                    let child = child.map_err(odraw_error)?;
                    if child.kind() == RecordKind::Bse {
                        match BlipStoreEntry::parse(child.data()) {
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

    /// Extract BLIP from BSE record, handling both embedded and delay-loaded cases.
    ///
    /// # Arguments
    /// * `bse` - Parsed BSE record
    /// * `record_data` - Raw BSE record data (without Escher header)
    /// * `delay_stream` - Optional data stream for delay-loaded BLIPs
    ///
    /// # Returns
    /// Owned BLIP data
    fn blip_from_bse(
        bse: &BlipStoreEntry<'_>,
        bse_record: &Record<'_>,
        delay_stream: Option<&[u8]>,
    ) -> Result<Blip<'static>> {
        let embedded_offset = 36usize
            .checked_add(usize::from(bse.name_len))
            .ok_or_else(|| {
                litchi_core::error::Error::ParseError("BSE name size overflow".into())
            })?;
        if embedded_offset > bse_record.data().len() {
            return Err(litchi_core::error::Error::ParseError(
                "BSE name extends beyond record data".into(),
            ));
        }
        let blip_data = if bse_record.data().len() == embedded_offset {
            if !bse.is_delay_loaded() {
                return Err(litchi_core::error::Error::ParseError(
                    "BSE has neither an embedded nor delay-loaded BLIP".into(),
                ));
            }
            let stream = delay_stream.ok_or_else(|| {
                litchi_core::error::Error::ParseError(
                    "BSE record is delay-loaded but no data stream was provided".into(),
                )
            })?;
            let offset = bse.offset as usize;
            if offset >= stream.len() {
                return Err(litchi_core::error::Error::ParseError(
                    "BSE delay stream offset is out of bounds".into(),
                ));
            }
            &stream[offset..]
        } else {
            &bse_record.data()[embedded_offset..]
        };

        if blip_data.len() < 8 {
            return Err(litchi_core::error::Error::ParseError(
                "Insufficient data for BLIP".into(),
            ));
        }
        let payload_length =
            u32::from_le_bytes([blip_data[4], blip_data[5], blip_data[6], blip_data[7]]);
        let record_length = usize::try_from(payload_length)
            .ok()
            .and_then(|length| 8usize.checked_add(length))
            .ok_or_else(|| {
                litchi_core::error::Error::ParseError("BLIP record size overflow".into())
            })?;
        if record_length > blip_data.len() || record_length as u64 != u64::from(bse.size) {
            return Err(litchi_core::error::Error::ParseError(
                "BSE size does not match its BLIP record".into(),
            ));
        }

        Blip::parse(&blip_data[..record_length]).map(|b| b.into_owned())
    }

    /// Extracts an image from one typed OfficeArt record.
    ///
    /// This method handles both BSE (0xF007) and BLIP type records directly.
    /// For BSE records, it extracts the embedded BLIP data.
    /// For BLIP records (0xF01A-0xF029), it parses the image directly.
    ///
    /// # Arguments
    /// * `record` - An OfficeArt record that is either BSE or a BLIP type
    ///
    /// # Returns
    /// ExtractedImage if the record contains valid image data
    pub fn extract_from_record(record: &Record<'_>) -> Result<ExtractedImage<'static>> {
        Self::extract_from_record_with_stream(record, None)
    }

    /// Extracts an image from one OfficeArt record with an optional data stream.
    ///
    /// This method handles both BSE (0xF007) and BLIP type records directly.
    /// For BSE records with delay-loaded BLIPs, the `delay_stream`
    /// parameter must be provided to locate the BLIP data.
    ///
    /// # Arguments
    /// * `record` - An OfficeArt record that is either BSE or a BLIP type
    /// * `delay_stream` - Optional main stream for delay-loaded BLIPs
    ///
    /// # Returns
    /// ExtractedImage if the record contains valid image data
    pub fn extract_from_record_with_stream(
        record: &Record<'_>,
        delay_stream: Option<&[u8]>,
    ) -> Result<ExtractedImage<'static>> {
        match record.kind() {
            // BLIP type records - parse directly
            RecordKind::BlipEmf
            | RecordKind::BlipWmf
            | RecordKind::BlipPict
            | RecordKind::BlipJpeg
            | RecordKind::BlipPng
            | RecordKind::BlipDib
            | RecordKind::BlipTiff => {
                // Reconstruct full BLIP record from the Escher record so
                // `Blip::parse` can decode it without depending on Escher
                // types from this crate.
                let mut full_data = Vec::with_capacity(8 + record.data().len());
                let ver_inst = (record.instance() << 4) | u16::from(record.version());
                full_data.extend_from_slice(&ver_inst.to_le_bytes());
                full_data.extend_from_slice(&record.raw_kind().to_le_bytes());
                full_data.extend_from_slice(&record.len().to_le_bytes());
                full_data.extend_from_slice(record.data());
                let blip = Blip::parse(&full_data).map(|b| b.into_owned())?;
                Ok(ExtractedImage::new(blip, None, 0))
            },

            // BSE record - extract embedded or delay-loaded BLIP
            RecordKind::Bse => {
                let bse = BlipStoreEntry::parse(record.data())?;
                let name = bse.name.as_ref().map(|n| n.to_string());
                let blip = Self::blip_from_bse(&bse, record, delay_stream)?;
                Ok(ExtractedImage::new(blip, name, 0))
            },
            _ => Err(litchi_core::error::Error::ParseError(format!(
                "Record type 0x{:04X} is not a supported image record",
                record.raw_kind()
            ))),
        }
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
        let parser = Parser::new(data);
        for record in parser.records() {
            let record = record.map_err(odraw_error)?;

            // Check if this is a BLIP record
            let is_blip = matches!(
                record.kind(),
                RecordKind::BlipEmf
                    | RecordKind::BlipWmf
                    | RecordKind::BlipPict
                    | RecordKind::BlipJpeg
                    | RecordKind::BlipPng
                    | RecordKind::BlipDib
                    | RecordKind::BlipTiff
            );

            if is_blip {
                // Need to reconstruct full BLIP record with header
                let mut full_data = Vec::with_capacity(8 + record.data().len());

                // Reconstruct header
                let ver_inst = (record.instance() << 4) | u16::from(record.version());
                full_data.extend_from_slice(&ver_inst.to_le_bytes());
                full_data.extend_from_slice(&record.raw_kind().to_le_bytes());
                full_data.extend_from_slice(&record.len().to_le_bytes());
                full_data.extend_from_slice(record.data());

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
            (0xF02A, "jpeg"),
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
    pub fn extract_from_container(container: &Container) -> Result<Vec<ExtractedImage<'static>>> {
        Self::extract_from_container_with_stream(container, None)
    }

    /// Extract images from a specific Escher container with optional data stream.
    ///
    /// # Arguments
    /// * `container` - The Escher container to extract from
    /// * `delay_stream` - Optional data stream for delay-loaded BLIPs
    pub fn extract_from_container_with_stream(
        container: &Container,
        delay_stream: Option<&[u8]>,
    ) -> Result<Vec<ExtractedImage<'static>>> {
        let mut images = Vec::new();
        let mut index = 0;

        // Recursively search for BLIP records
        Self::extract_from_container_recursive_with_stream(
            container,
            &mut images,
            &mut index,
            delay_stream,
        )?;

        Ok(images)
    }

    /// Recursively extract BLIPs from a container with optional data stream.
    fn extract_from_container_recursive_with_stream(
        container: &Container,
        images: &mut Vec<ExtractedImage<'static>>,
        index: &mut usize,
        delay_stream: Option<&[u8]>,
    ) -> Result<()> {
        // Check if this container has BLIP records
        for child_result in container.children() {
            let child = child_result.map_err(odraw_error)?;

            if child.kind().is_blip() {
                // Reconstruct full BLIP record
                let mut full_data = Vec::with_capacity(8 + child.data().len());
                let ver_inst = (child.instance() << 4) | u16::from(child.version());
                full_data.extend_from_slice(&ver_inst.to_le_bytes());
                full_data.extend_from_slice(&child.raw_kind().to_le_bytes());
                full_data.extend_from_slice(&child.len().to_le_bytes());
                full_data.extend_from_slice(child.data());

                if let Ok(blip) = Blip::parse(&full_data) {
                    // Convert to owned to avoid lifetime issues
                    let owned_blip = blip.into_owned();
                    images.push(ExtractedImage::new(owned_blip, None, *index));
                    *index += 1;
                }
            } else if child.is_container() {
                // Recurse into child containers
                let child_container = Container::try_new(child).map_err(odraw_error)?;
                Self::extract_from_container_recursive_with_stream(
                    &child_container,
                    images,
                    index,
                    delay_stream,
                )?;
            } else if child.kind() == RecordKind::Bse {
                // BSE records can contain embedded or delay-loaded BLIP data
                match BlipStoreEntry::parse(child.data()) {
                    Ok(bse) => {
                        let name = bse.name.as_ref().map(|n| n.to_string());
                        match Self::blip_from_bse(&bse, &child, delay_stream) {
                            Ok(blip) => {
                                images.push(ExtractedImage::new(blip, name, *index));
                                *index += 1;
                            },
                            Err(_) => {
                                continue;
                            },
                        }
                    },
                    Err(_) => {
                        continue;
                    },
                }
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
pub mod ppt {
    use super::*;
    use crate::OleFile;
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
pub mod doc {
    use super::*;
    use crate::OleFile;
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
            0xA0, 0x46, // version=0, instance=0x46A
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

    fn png_blip() -> Vec<u8> {
        let mut payload = vec![0; 16];
        payload.push(0xff);
        payload.extend_from_slice(&[1, 2, 3]);
        let mut record = Vec::new();
        record.extend_from_slice(&(0x6e0u16 << 4).to_le_bytes());
        record.extend_from_slice(&0xf01eu16.to_le_bytes());
        record.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        record.extend_from_slice(&payload);
        record
    }

    fn fbse(blip: Option<&[u8]>, offset: u32, name: &[u8]) -> Vec<u8> {
        let mut payload = vec![0x06, 0x06];
        payload.extend_from_slice(&[0; 16]);
        payload.extend_from_slice(&0xffu16.to_le_bytes());
        payload.extend_from_slice(&(blip.map_or(28, <[u8]>::len) as u32).to_le_bytes());
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&offset.to_le_bytes());
        payload.push(0);
        payload.push(name.len() as u8);
        payload.extend_from_slice(&[0, 0]);
        payload.extend_from_slice(name);
        if let Some(blip) = blip {
            payload.extend_from_slice(blip);
        }
        let mut record = Vec::new();
        record.extend_from_slice(&0x62u16.to_le_bytes());
        record.extend_from_slice(&0xf007u16.to_le_bytes());
        record.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        record.extend_from_slice(&payload);
        record
    }

    #[test]
    fn extracts_named_embedded_and_offset_zero_bse_blips() {
        let blip = png_blip();
        let embedded = fbse(Some(&blip), u32::MAX, &[b'A', 0, 0, 0]);
        let (record, _) = Record::parse(&embedded, 0).unwrap();
        let image = ImageExtractor::extract_from_record(&record).unwrap();
        assert_eq!(image.name.as_deref(), Some("A"));
        assert_eq!(image.blip.picture_data(), [1, 2, 3]);

        let delayed = fbse(None, 0, &[]);
        let (record, _) = Record::parse(&delayed, 0).unwrap();
        let image = ImageExtractor::extract_from_record_with_stream(&record, Some(&blip)).unwrap();
        assert_eq!(image.blip.picture_data(), [1, 2, 3]);
    }
}
