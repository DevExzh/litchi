use super::model::{LazyPictures, Presentation, PresentationData};
#[cfg(not(feature = "encryption"))]
use crate::current_user::CurrentUser;
#[cfg(feature = "encryption")]
use crate::encryption::{decrypt_pictures, decrypt_powerpoint_document};
use crate::package::{Error, OpenOptions, Result};
use crate::parsers::RecordParser;
#[cfg(feature = "encryption")]
use crate::persist::PersistMapping;
use crate::slide::SlideDirectory;
use litchi_cfb::{OleFile, SharedOleFile};
use litchi_ole_common::property_set::{PropertySetReader as _, SharedPropertySetReader};
use std::io::{Read, Seek};
use std::sync::Arc;

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
        source_limits: crate::SourceLimits,
    ) -> Result<Self> {
        // Resolve and charge every declared stream before materializing any
        // payload. This covers metadata and opaque streams that are reachable
        // through the package's raw-property APIs as well as native streams.
        crate::package::preflight_ole_file(ole, source_limits)?;
        // Resolve the native streams used by the presentation projection.
        let document = locate_stream(
            ole,
            &DOCUMENT_PATHS,
            source_limits.max_input_bytes,
            "PowerPoint Document",
        )?
        .ok_or_else(|| Error::InvalidFormat("PowerPoint Document stream not found".to_string()))?;
        let current_user = locate_stream(
            ole,
            &CURRENT_USER_PATHS,
            source_limits.max_input_bytes,
            "Current User",
        )?;
        let pictures = locate_stream(
            ole,
            &PICTURES_PATHS,
            source_limits.max_input_bytes,
            "Pictures",
        )?;
        let aggregate_bytes = document
            .1
            .checked_add(current_user.map_or(0, |value| value.1))
            .and_then(|value| value.checked_add(pictures.map_or(0, |entry| entry.1)))
            .ok_or(Error::Limit {
                kind: crate::SourceLimitKind::AggregateBytes,
                observed: usize::MAX,
                maximum: source_limits.max_aggregate_input_bytes,
            })?;
        if aggregate_bytes > source_limits.max_aggregate_input_bytes {
            return Err(Error::Limit {
                kind: crate::SourceLimitKind::AggregateBytes,
                observed: aggregate_bytes,
                maximum: source_limits.max_aggregate_input_bytes,
            });
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
            source_limits,
        )?;
        #[cfg(not(feature = "encryption"))]
        {
            let _ = options;
            if let Some(parsed_current_user) = current_user_data
                .as_deref()
                .map(|data| CurrentUser::parse(data, &source_limits))
                .transpose()?
                && parsed_current_user.is_encrypted()
            {
                return Err(Error::UnsupportedEncryption(
                    crate::package::EncryptionKind::CryptoApi,
                ));
            }
        }

        // Parse document structure. The stream is retained by the presentation
        // anyway, so the record tree borrows its payloads from it.
        //
        // `tolerance` is the open-time aggregation point for repaired
        // non-structural defects (see `crate::leniency`).
        let mut tolerance = crate::leniency::ToleranceReport::default();
        // Container repairs come first: the compound-file reader applied them
        // while the OLE file was opened, before any record was read.
        tolerance.append_container(ole.tolerance_report());
        let powerpoint_document = Arc::new(powerpoint_document);
        let mut parser = RecordParser::new();
        #[cfg(feature = "encryption")]
        if let Some(encryption) = &encrypted {
            parser.parse_shared_document_at_offsets_with_limits(
                &powerpoint_document,
                &encryption.live_offsets,
                source_limits,
            )?;
        } else {
            parser.parse_shared_document_with_limits(&powerpoint_document, source_limits)?;
        }
        #[cfg(not(feature = "encryption"))]
        parser.parse_shared_document_with_limits(&powerpoint_document, source_limits)?;

        // The live UserEditAtom, recovered by a stream scan when the Current
        // User stream is missing or its edit offset is unusable (F5.6).
        let head_user_edit = crate::slide::directory::locate_head_user_edit(
            &powerpoint_document,
            current_user_data.as_deref(),
            source_limits,
            &mut tolerance,
        )?;
        #[cfg(feature = "encryption")]
        let persist_mapping = if let Some(encryption) = &encrypted {
            let mut mapping = PersistMapping::new();
            for &(persist_id, offset) in &encryption.mappings {
                mapping.add_mapping(persist_id, offset);
            }
            mapping
        } else {
            crate::embedded::object::Editor::inspect_live_mapping_at(
                &powerpoint_document,
                head_user_edit,
                options.leniency,
                &mut tolerance,
            )?
        };
        #[cfg(not(feature = "encryption"))]
        let persist_mapping = crate::embedded::object::Editor::inspect_live_mapping_at(
            &powerpoint_document,
            head_user_edit,
            options.leniency,
            &mut tolerance,
        )?;
        let slide_directory = SlideDirectory::build_from_head(
            &powerpoint_document,
            head_user_edit,
            &persist_mapping,
            source_limits,
            &mut tolerance,
        )?;

        // Try to read Pictures stream for image extraction
        let pictures_data = if let Some((index, _)) = pictures {
            #[cfg(feature = "encryption")]
            {
                let mut stream = ole.open_stream(PICTURES_PATHS[index])?;
                if let Some(encryption) = &encrypted {
                    decrypt_pictures(&mut stream, &encryption.crypto, source_limits)?;
                }
                Some(stream)
            }
            #[cfg(not(feature = "encryption"))]
            {
                Some(ole.open_stream(PICTURES_PATHS[index])?)
            }
        } else {
            None
        };

        // The metadata streams' repairs follow every record-level repair.
        tolerance.append_property_sets(|path| ole.property_set_stream(path));

        #[cfg(feature = "encryption")]
        let encryption_profile = encrypted.as_ref().map(|value| value.profile);
        Ok(Self::from_data(PresentationData {
            powerpoint_document,
            parser,
            persist_mapping,
            slide_directory,
            pictures_data,
            pictures_source: None,
            source_limits,
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: options.vba_project_limits_value(),
            tolerance,
            source: None,
            #[cfg(feature = "encryption")]
            encryption: encryption_profile,
            shape_memo: crate::slide::ShapeMemo::default(),
        }))
    }

    /// Build a presentation from one validated positional CFB index.
    ///
    /// The required document and Current User streams follow the established
    /// eager semantic path.  A present `Pictures` stream is retained only as
    /// a descriptor; its payload is loaded by the image query methods.
    pub(crate) fn from_shared_with_options(
        shared: Arc<SharedOleFile>,
        options: OpenOptions<'_>,
        source_limits: crate::SourceLimits,
    ) -> Result<Self> {
        // Fence the positional source before any metadata or limit decision;
        // a stale source must never be reported as a fresh limit failure.
        shared.source_version()?;
        crate::package::preflight_shared_ole_file(&shared, source_limits)?;
        let package_bytes = usize::try_from(shared.file_size()).map_err(|_error| Error::Limit {
            kind: crate::SourceLimitKind::InputBytes,
            observed: usize::MAX,
            maximum: source_limits.max_package_bytes,
        })?;
        if package_bytes > source_limits.max_package_bytes {
            return Err(Error::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed: package_bytes,
                maximum: source_limits.max_package_bytes,
            });
        }

        let document = locate_shared_stream(
            &shared,
            &DOCUMENT_PATHS,
            source_limits.max_input_bytes,
            "PowerPoint Document",
        )?
        .ok_or_else(|| Error::InvalidFormat("PowerPoint Document stream not found".to_string()))?;
        let current_user = locate_shared_stream(
            &shared,
            &CURRENT_USER_PATHS,
            source_limits.max_input_bytes,
            "Current User",
        )?;
        let pictures = locate_shared_stream(
            &shared,
            &PICTURES_PATHS,
            source_limits.max_input_bytes,
            "Pictures",
        )?;
        let aggregate_bytes = document
            .1
            .checked_add(current_user.map_or(0, |value| value.1))
            .and_then(|value| value.checked_add(pictures.map_or(0, |entry| entry.1)))
            .ok_or(Error::Limit {
                kind: crate::SourceLimitKind::AggregateBytes,
                observed: usize::MAX,
                maximum: source_limits.max_aggregate_input_bytes,
            })?;
        if aggregate_bytes > source_limits.max_aggregate_input_bytes {
            return Err(Error::Limit {
                kind: crate::SourceLimitKind::AggregateBytes,
                observed: aggregate_bytes,
                maximum: source_limits.max_aggregate_input_bytes,
            });
        }

        #[cfg(feature = "encryption")]
        let mut powerpoint_document = shared.open_stream(document.0)?;
        #[cfg(not(feature = "encryption"))]
        let powerpoint_document = shared.open_stream(document.0)?;
        let current_user_data = current_user
            .map(|(path, _)| shared.open_stream(path))
            .transpose()?;
        #[cfg(feature = "encryption")]
        let encrypted = decrypt_powerpoint_document(
            &mut powerpoint_document,
            current_user_data.as_deref(),
            options.password,
            source_limits,
        )?;
        #[cfg(not(feature = "encryption"))]
        {
            let _ = options;
            if let Some(parsed_current_user) = current_user_data
                .as_deref()
                .map(|data| CurrentUser::parse(data, &source_limits))
                .transpose()?
                && parsed_current_user.is_encrypted()
            {
                return Err(Error::UnsupportedEncryption(
                    crate::package::EncryptionKind::CryptoApi,
                ));
            }
        }

        // `tolerance` is the open-time aggregation point for repaired
        // non-structural defects (see `crate::leniency`).
        let mut tolerance = crate::leniency::ToleranceReport::default();
        // Container repairs come first: the compound-file reader applied them
        // while the OLE file was opened, before any record was read.
        tolerance.append_container(shared.tolerance_report());
        let powerpoint_document = Arc::new(powerpoint_document);
        let mut parser = RecordParser::new();
        #[cfg(feature = "encryption")]
        if let Some(encryption) = &encrypted {
            parser.parse_shared_document_at_offsets_with_limits(
                &powerpoint_document,
                &encryption.live_offsets,
                source_limits,
            )?;
        } else {
            parser.parse_shared_document_with_limits(&powerpoint_document, source_limits)?;
        }
        #[cfg(not(feature = "encryption"))]
        parser.parse_shared_document_with_limits(&powerpoint_document, source_limits)?;

        // The live UserEditAtom, recovered by a stream scan when the Current
        // User stream is missing or its edit offset is unusable (F5.6).
        let head_user_edit = crate::slide::directory::locate_head_user_edit(
            &powerpoint_document,
            current_user_data.as_deref(),
            source_limits,
            &mut tolerance,
        )?;
        #[cfg(feature = "encryption")]
        let persist_mapping = if let Some(encryption) = &encrypted {
            let mut mapping = PersistMapping::new();
            for &(persist_id, offset) in &encryption.mappings {
                mapping.add_mapping(persist_id, offset);
            }
            mapping
        } else {
            crate::embedded::object::Editor::inspect_live_mapping_at(
                &powerpoint_document,
                head_user_edit,
                options.leniency,
                &mut tolerance,
            )?
        };
        #[cfg(not(feature = "encryption"))]
        let persist_mapping = crate::embedded::object::Editor::inspect_live_mapping_at(
            &powerpoint_document,
            head_user_edit,
            options.leniency,
            &mut tolerance,
        )?;
        let slide_directory = SlideDirectory::build_from_head(
            &powerpoint_document,
            head_user_edit,
            &persist_mapping,
            source_limits,
            &mut tolerance,
        )?;

        #[cfg(feature = "encryption")]
        let encryption_profile = encrypted.as_ref().map(|value| value.profile);
        #[cfg(feature = "encryption")]
        let pictures_crypto = encrypted.map(|value| value.crypto);
        let pictures_source = pictures.map(|(path, length)| {
            LazyPictures::new(
                Arc::clone(&shared),
                path,
                length,
                source_limits,
                #[cfg(feature = "encryption")]
                pictures_crypto,
            )
        });
        // The metadata streams' repairs follow every record-level repair.
        tolerance.append_property_sets(|path| {
            <SharedOleFile as SharedPropertySetReader>::property_set_stream(&shared, path)
        });
        shared.source_version()?;

        Ok(Self::from_data(PresentationData {
            powerpoint_document,
            parser,
            persist_mapping,
            slide_directory,
            pictures_data: None,
            pictures_source,
            source_limits,
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: options.vba_project_limits_value(),
            tolerance,
            source: Some(shared),
            #[cfg(feature = "encryption")]
            encryption: encryption_profile,
            shape_memo: crate::slide::ShapeMemo::default(),
        }))
    }
}

fn locate_shared_stream(
    shared: &SharedOleFile,
    paths: &[&'static [&'static str]],
    max_bytes: usize,
    _label: &'static str,
) -> Result<Option<(&'static [&'static str], usize)>> {
    for path in paths {
        match shared.stream_len(path) {
            Ok(size) => {
                let size_bytes = usize::try_from(size).map_err(|_err| Error::Limit {
                    kind: crate::SourceLimitKind::EntryBytes,
                    observed: usize::MAX,
                    maximum: max_bytes,
                })?;
                if size_bytes > max_bytes {
                    return Err(Error::Limit {
                        kind: crate::SourceLimitKind::EntryBytes,
                        observed: size_bytes,
                        maximum: max_bytes,
                    });
                }
                return Ok(Some((*path, size_bytes)));
            },
            Err(litchi_cfb::OleError::StreamNotFound) => {},
            Err(error) => return Err(error.into()),
        }
    }
    Ok(None)
}

fn locate_stream<R: Read + Seek>(
    ole: &OleFile<R>,
    paths: &[&[&str]],
    max_bytes: usize,
    _label: &'static str,
) -> Result<Option<(usize, usize)>> {
    for (index, path) in paths.iter().enumerate() {
        match ole.stream_len(path) {
            Ok(size) => {
                let size_bytes = usize::try_from(size).map_err(|_err| Error::Limit {
                    kind: crate::SourceLimitKind::EntryBytes,
                    observed: usize::MAX,
                    maximum: max_bytes,
                })?;
                if size_bytes > max_bytes {
                    return Err(Error::Limit {
                        kind: crate::SourceLimitKind::EntryBytes,
                        observed: size_bytes,
                        maximum: max_bytes,
                    });
                }
                return Ok(Some((index, size_bytes)));
            },
            Err(litchi_cfb::OleError::StreamNotFound) => {},
            Err(error) => return Err(error.into()),
        }
    }
    Ok(None)
}
