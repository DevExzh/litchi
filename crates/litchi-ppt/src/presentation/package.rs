use super::model::Presentation;
use crate::encryption::{decrypt_pictures, decrypt_powerpoint_document};
use crate::package::{Error, OpenOptions, Result};
use crate::parsers::RecordParser;
use crate::persist::PersistMapping;
use crate::slide::SlideDirectory;
use litchi_cfb::OleFile;
use std::io::{Read, Seek};

impl Presentation {
    /// Create a new Presentation from an OLE file.
    pub(crate) fn from_ole<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Self> {
        Self::from_ole_with_options(ole, OpenOptions::default())
    }

    /// Create a new Presentation from an OLE file with password-to-open options.
    pub(crate) fn from_ole_with_options<R: Read + Seek>(
        ole: &mut OleFile<R>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        // Read the PowerPoint Document stream
        let mut powerpoint_document = Self::read_powerpoint_document(ole)?;
        let current_user_data = ole
            .open_stream(&["Current User"])
            .or_else(|_| ole.open_stream(&["PP97_DUALSTORAGE", "Current User"]))
            .ok();
        let encrypted = decrypt_powerpoint_document(
            &mut powerpoint_document,
            current_user_data.as_deref(),
            options.password,
        )?;

        // Parse document structure
        let mut parser = RecordParser::new();
        if let Some(encrypted) = &encrypted {
            parser.parse_document_at_offsets(&powerpoint_document, &encrypted.live_offsets)?;
        } else {
            parser.parse_document(&powerpoint_document)?;
        }

        // Build persist mapping for slide lookup (collect all records recursively)
        // Use zero-copy reference collection to avoid cloning all record data
        let all_records_ref = parser.find_records_ref();
        let mut persist_mapping = PersistMapping::build_from_records_ref(&all_records_ref);
        if let Some(encrypted) = &encrypted {
            persist_mapping = PersistMapping::new();
            for &(persist_id, offset) in &encrypted.mappings {
                persist_mapping.add_mapping(persist_id, offset);
            }
        }
        let current_user_data = current_user_data
            .as_deref()
            .ok_or_else(|| Error::StreamNotFound("Current User".to_string()))?;
        let slide_directory =
            SlideDirectory::build(&powerpoint_document, current_user_data, &persist_mapping)?;

        // Try to read Pictures stream for image extraction
        let pictures_data = if let Ok(mut pictures) = ole.open_stream(&["Pictures"]) {
            if let Some(encrypted) = &encrypted {
                decrypt_pictures(&mut pictures, &encrypted.crypto)?;
            }
            Some(pictures)
        } else {
            None
        };

        Ok(Self {
            powerpoint_document,
            parser,
            persist_mapping,
            slide_directory,
            pictures_data,
        })
    }

    /// Read the PowerPoint Document stream from OLE file.
    fn read_powerpoint_document<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Vec<u8>> {
        // Try primary location
        if let Ok(data) = ole.open_stream(&["PowerPoint Document"]) {
            return Ok(data);
        }

        // Try alternate location
        if let Ok(data) = ole.open_stream(&["PP97_DUALSTORAGE", "PowerPoint Document"]) {
            return Ok(data);
        }

        Err(Error::InvalidFormat(
            "PowerPoint Document stream not found".to_string(),
        ))
    }
}
