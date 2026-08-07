use super::model::Presentation;
#[cfg(not(feature = "encryption"))]
use crate::current_user::CurrentUser;
#[cfg(feature = "encryption")]
use crate::encryption::{decrypt_pictures, decrypt_powerpoint_document};
use crate::package::{Error, OpenOptions, Result};
use crate::parsers::RecordParser;
#[cfg(feature = "encryption")]
use crate::persist::PersistMapping;
use crate::slide::SlideDirectory;
use litchi_cfb::OleFile;
use std::io::{Read, Seek};

const DOCUMENT_PATHS: [&[&str]; 2] = [
    &["PowerPoint Document"],
    &["PP97_DUALSTORAGE", "PowerPoint Document"],
];
const CURRENT_USER_PATHS: [&[&str]; 2] = [&["Current User"], &["PP97_DUALSTORAGE", "Current User"]];
const PICTURES_PATHS: [&[&str]; 1] = [&["Pictures"]];

impl Presentation {
    /// Create a new Presentation from an OLE file with password-to-open options.
    pub(crate) fn from_ole_with_options<R: Read + Seek>(
        ole: &mut OleFile<R>,
        options: OpenOptions<'_>,
        record_limits: crate::RecordLimits,
    ) -> Result<Self> {
        // Resolve and charge all hostile stream sizes before materializing any payload.
        let document = locate_stream(
            ole,
            &DOCUMENT_PATHS,
            record_limits.max_input_bytes,
            "PowerPoint Document",
        )?
        .ok_or_else(|| Error::InvalidFormat("PowerPoint Document stream not found".to_string()))?;
        let current_user = locate_stream(
            ole,
            &CURRENT_USER_PATHS,
            record_limits.max_input_bytes,
            "Current User",
        )?;
        let pictures = locate_stream(
            ole,
            &PICTURES_PATHS,
            record_limits.max_input_bytes,
            "Pictures",
        )?;
        let aggregate_bytes = document
            .1
            .checked_add(current_user.map_or(0, |value| value.1))
            .and_then(|value| value.checked_add(pictures.map_or(0, |value| value.1)))
            .ok_or_else(|| Error::ResourceLimit("PPT stream-size sum overflow".to_string()))?;
        if aggregate_bytes > record_limits.max_aggregate_input_bytes {
            return Err(Error::ResourceLimit(format!(
                "PPT aggregate stream size {aggregate_bytes} exceeds limit {}",
                record_limits.max_aggregate_input_bytes
            )));
        }

        #[cfg(feature = "encryption")]
        let mut powerpoint_document = ole.open_stream(DOCUMENT_PATHS[document.0])?;
        #[cfg(not(feature = "encryption"))]
        let powerpoint_document = ole.open_stream(DOCUMENT_PATHS[document.0])?;
        let current_user_data = current_user
            .map(|(index, _)| ole.open_stream(CURRENT_USER_PATHS[index]))
            .transpose()?;
        #[cfg(feature = "encryption")]
        let encrypted = decrypt_powerpoint_document(
            &mut powerpoint_document,
            current_user_data.as_deref(),
            options.password,
            record_limits,
        )?;
        #[cfg(not(feature = "encryption"))]
        {
            let _ = options;
            if let Some(current_user) = current_user_data
                .as_deref()
                .map(CurrentUser::parse)
                .transpose()?
                && current_user.is_encrypted()
            {
                return Err(Error::UnsupportedEncryption(
                    crate::package::EncryptionKind::CryptoApi,
                ));
            }
        }

        // Parse document structure
        let mut parser = RecordParser::new();
        #[cfg(feature = "encryption")]
        if let Some(encrypted) = &encrypted {
            parser.parse_document_at_offsets_with_limits(
                &powerpoint_document,
                &encrypted.live_offsets,
                record_limits,
            )?;
        } else {
            parser.parse_document_with_limits(&powerpoint_document, record_limits)?;
        }
        #[cfg(not(feature = "encryption"))]
        parser.parse_document_with_limits(&powerpoint_document, record_limits)?;

        let current_user_data = current_user_data
            .as_deref()
            .ok_or_else(|| Error::StreamNotFound("Current User".to_string()))?;
        #[cfg(feature = "encryption")]
        let persist_mapping = if let Some(encrypted) = &encrypted {
            let mut mapping = PersistMapping::new();
            for &(persist_id, offset) in &encrypted.mappings {
                mapping.add_mapping(persist_id, offset);
            }
            mapping
        } else {
            crate::embedded::object::Editor::inspect_live_mapping(
                &powerpoint_document,
                current_user_data,
            )?
        };
        #[cfg(not(feature = "encryption"))]
        let persist_mapping = crate::embedded::object::Editor::inspect_live_mapping(
            &powerpoint_document,
            current_user_data,
        )?;
        let slide_directory = SlideDirectory::build_with_limits(
            &powerpoint_document,
            current_user_data,
            &persist_mapping,
            record_limits,
        )?;

        // Try to read Pictures stream for image extraction
        let pictures_data = if let Some((index, _)) = pictures {
            let pictures = ole.open_stream(PICTURES_PATHS[index])?;
            #[cfg(feature = "encryption")]
            let mut pictures = pictures;
            #[cfg(feature = "encryption")]
            if let Some(encrypted) = &encrypted {
                decrypt_pictures(&mut pictures, &encrypted.crypto, record_limits)?;
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
            record_limits,
        })
    }
}

fn locate_stream<R: Read + Seek>(
    ole: &OleFile<R>,
    paths: &[&[&str]],
    max_bytes: usize,
    label: &'static str,
) -> Result<Option<(usize, usize)>> {
    for (index, path) in paths.iter().enumerate() {
        match ole.stream_len(path) {
            Ok(size) => {
                let size = usize::try_from(size).map_err(|_| {
                    Error::ResourceLimit(format!("{label} stream size exceeds this platform"))
                })?;
                if size > max_bytes {
                    return Err(Error::ResourceLimit(format!(
                        "{label} stream size {size} exceeds limit {max_bytes}"
                    )));
                }
                return Ok(Some((index, size)));
            },
            Err(litchi_cfb::OleError::StreamNotFound) => {},
            Err(error) => return Err(error.into()),
        }
    }
    Ok(None)
}
