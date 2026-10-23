//! The path route of [`crate::open`].
//!
//! On `unix` and `windows` the file is opened once as a positional source,
//! pinned to the revision it had when opened, and every probe and owner reads
//! that source; a change underneath is reported, not silently mixed in. On
//! other targets the file is read once and takes the byte route.

use std::path::Path;
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "xlsx",
    feature = "pptx",
    feature = "rtf",
    feature = "xlsb",
    feature = "ods",
    feature = "odt",
    feature = "odp",
    odf_objects,
    feature = "iwork-adapter",
    all(
        any(unix, windows),
        any(
            feature = "doc",
            feature = "ppt",
            feature = "xls",
            feature = "encryption"
        )
    )
))]
use std::sync::Arc;

use crate::detect::tail::ZipTail;
use crate::error::Result;
use crate::{File, OpenOptions};

/// Resolve a path origin once, at open time, so later saves do not depend on
/// the caller's current working directory. This deliberately makes no
/// filename or extension decision and does not canonicalize through the
/// filesystem: a path's symlink spelling remains the path the caller opened.
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "xlsx",
    feature = "pptx",
    feature = "rtf",
    feature = "xlsb",
    feature = "ods",
    feature = "odt",
    feature = "odp",
    odf_objects,
    feature = "iwork-adapter",
    all(
        any(unix, windows),
        any(
            feature = "doc",
            feature = "ppt",
            feature = "xls",
            feature = "encryption"
        )
    )
))]
pub(crate) fn stable_origin(path: &Path) -> Result<Arc<Path>> {
    let absolute = if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir()?.join(path)
    };
    Ok(Arc::from(absolute.into_boxed_path()))
}

/// Open `path` through the route the content selects.
#[cfg(any(unix, windows))]
pub(super) fn open<T: ZipTail>(path: &Path, options: &OpenOptions, tail: &T) -> Result<File> {
    use std::sync::Arc;

    use litchi_core::ReadAt;
    use litchi_core::detection::simd_utils::signature_matches;
    use litchi_core::detection::utils::{OLE2_SIGNATURE, ZIP_SIGNATURE};

    use crate::detect::detected::ensure_path_source_current;

    let source: Arc<dyn ReadAt> = Arc::new(litchi_core::FileSource::open(path)?);
    let version = source.version()?;
    let length = source.len()?;
    let ceiling = super::input_ceiling(options);
    super::check_input_len(length, ceiling)?;

    let mut header = [0_u8; 8];
    let read = read_prefix(source.as_ref(), &mut header)?;
    ensure_path_source_current(source.as_ref(), version)?;
    let header = &header[..read];
    let ole2 = header.len() == OLE2_SIGNATURE.len() && signature_matches(header, OLE2_SIGNATURE);
    let zip = header.len() >= ZIP_SIGNATURE.len() && signature_matches(header, ZIP_SIGNATURE);

    // An OOXML file name selects the OPC read policy. A ZIP applies it once
    // ordinary ODF has had its turn; any other input is held to its input
    // ceiling here and refused, as the facade refused it, unless it is a
    // compound file: an encrypted OOXML package arrives in one, and the
    // compound-file route refuses every other host under this name.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    if !zip && super::ooxml_format_for_extension(path).is_some() {
        let limits = options.ooxml_limits.unwrap_or_default();
        if length > limits.max_input_bytes() {
            return Err(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::InputBytes,
                observed: length,
                maximum: limits.max_input_bytes(),
            }
            .into());
        }
        if !ole2 {
            return Err(super::ooxml_suffix_mismatch());
        }
    }

    if options.source_mode == crate::SourceMode::Eager || !(ole2 || zip) {
        let bytes = super::read_whole_source(source.as_ref(), version, ceiling)?;
        let file = super::bytes::open(
            bytes,
            options,
            tail,
            super::ooxml_format_for_extension(path),
        )?;
        #[cfg(any(
            feature = "doc",
            feature = "docx",
            feature = "ppt",
            feature = "xlsx",
            feature = "pptx",
            feature = "xls",
            feature = "xlsb",
            feature = "rtf",
            feature = "ods",
            feature = "odp",
            feature = "odt",
            odf_objects,
            feature = "iwork-adapter"
        ))]
        let file = if matches!(
            file.format(),
            Some(
                litchi_core::FileFormat::Xls
                    | litchi_core::FileFormat::Doc
                    | litchi_core::FileFormat::Docx
                    | litchi_core::FileFormat::Ppt
                    | litchi_core::FileFormat::Xlsx
                    | litchi_core::FileFormat::Pptx
                    | litchi_core::FileFormat::Rtf
                    | litchi_core::FileFormat::Xlsb
                    | litchi_core::FileFormat::Ods
                    | litchi_core::FileFormat::Odp
                    | litchi_core::FileFormat::Pages
                    | litchi_core::FileFormat::Keynote
                    | litchi_core::FileFormat::Numbers
            )
        ) || matches!(&file, File::Locked(_))
        {
            super::attach_origin(file, stable_origin(path)?)
        } else {
            file
        };
        return Ok(file);
    }

    if ole2 {
        #[cfg(any(
            feature = "doc",
            feature = "ppt",
            feature = "xls",
            feature = "encryption"
        ))]
        {
            let file = super::cfb::open_path(path, source, version, options)?;
            #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
            let file = if matches!(
                file.format(),
                Some(
                    litchi_core::FileFormat::Doc
                        | litchi_core::FileFormat::Ppt
                        | litchi_core::FileFormat::Xls,
                )
            ) {
                super::attach_origin(file, stable_origin(path)?)
            } else {
                file
            };
            return Ok(file);
        }
        #[cfg(not(any(
            feature = "doc",
            feature = "ppt",
            feature = "xls",
            feature = "encryption"
        )))]
        return Err(crate::Error::unsupported_format(None));
    }
    zip::open(path, source, version, length, options, tail)
}

/// Open `path` on a target without positional file sources: read it once
/// and take the byte route.
#[cfg(not(any(unix, windows)))]
pub(super) fn open<T: ZipTail>(path: &Path, options: &OpenOptions, tail: &T) -> Result<File> {
    use std::io::Read;

    let mut file = std::fs::File::open(path)?;
    let length = file.metadata()?.len();
    let ceiling = super::input_ceiling(options);
    super::check_input_len(length, ceiling)?;
    let capacity = usize::try_from(length).map_err(|_| {
        crate::Error::from(litchi_core::Error::ResourceLimit(
            litchi_core::ResourceLimit {
                resource: litchi_core::Resource::InputBytes,
                observed: length,
                limit: u64::try_from(usize::MAX).unwrap_or(u64::MAX),
                scope: std::sync::Arc::from("unified open input"),
            },
        ))
    })?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(capacity)
        .map_err(|source| litchi_core::Error::Allocation {
            resource: "unified open input bytes",
            source,
        })?;
    (&mut file)
        .take(ceiling.saturating_add(1))
        .read_to_end(&mut bytes)?;
    super::check_input_len(u64::try_from(bytes.len()).unwrap_or(u64::MAX), ceiling)?;
    let opened = super::bytes::open(
        bytes,
        options,
        tail,
        super::ooxml_format_for_extension(path),
    )?;
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "xlsx",
        feature = "pptx",
        feature = "xls",
        feature = "xlsb",
        feature = "rtf",
        feature = "ods",
        feature = "odp",
        feature = "odt",
        odf_objects,
        feature = "iwork-adapter"
    ))]
    let opened = if matches!(
        opened.format(),
        Some(
            litchi_core::FileFormat::Xls
                | litchi_core::FileFormat::Doc
                | litchi_core::FileFormat::Docx
                | litchi_core::FileFormat::Ppt
                | litchi_core::FileFormat::Xlsx
                | litchi_core::FileFormat::Pptx
                | litchi_core::FileFormat::Rtf
                | litchi_core::FileFormat::Xlsb
                | litchi_core::FileFormat::Ods
                | litchi_core::FileFormat::Odp
                | litchi_core::FileFormat::Oth
                | litchi_core::FileFormat::Odm
                | litchi_core::FileFormat::Odg
                | litchi_core::FileFormat::Odc
                | litchi_core::FileFormat::Odi
                | litchi_core::FileFormat::Odb
                | litchi_core::FileFormat::Odf
                | litchi_core::FileFormat::Pages
                | litchi_core::FileFormat::Keynote
                | litchi_core::FileFormat::Numbers
        )
    ) || matches!(&opened, File::Locked(_))
    {
        super::attach_origin(opened, stable_origin(path)?)
    } else {
        opened
    };
    Ok(opened)
}

/// Fill `header` from the start of `source`, stopping at its end.
#[cfg(any(unix, windows))]
fn read_prefix(source: &dyn litchi_core::ReadAt, header: &mut [u8]) -> Result<usize> {
    let mut filled = 0;
    while filled < header.len() {
        let read = source.read_at(filled as u64, &mut header[filled..])?;
        if read == 0 {
            break;
        }
        filled += read;
    }
    Ok(filled)
}

/// The ZIP arm of the positional route.
#[cfg(any(unix, windows))]
mod zip {
    use std::path::Path;
    use std::sync::Arc;

    use litchi_core::{FileFormat, ReadAt, SourceVersion};

    #[cfg(any(
        feature = "docx",
        feature = "pptx",
        feature = "xlsx",
        feature = "xlsb",
        feature = "odt",
        feature = "ods",
        feature = "odp",
        feature = "odt",
        odf_objects
    ))]
    use crate::detect::detected::ensure_path_source_current;
    use crate::detect::tail::ZipTail;
    use crate::error::Result;
    use crate::{File, OpenOptions};

    /// Classify a positional ZIP and adopt the owner its format keeps.
    #[allow(
        clippy::too_many_lines,
        reason = "one precedence chain, kept in one place so its order is reviewable"
    )]
    pub(super) fn open<T: ZipTail>(
        path: &Path,
        source: Arc<dyn ReadAt>,
        version: SourceVersion,
        length: u64,
        options: &OpenOptions,
        tail: &T,
    ) -> Result<File> {
        let detection_limits = options.detection_limits();
        // 1. An ordinary ODF package: an ODF `mimetype` and no OPC catalog.
        #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
        let odf = {
            let mime = litchi_odf_common::detect::packaged_mime_read_at(
                source.as_ref(),
                litchi_odf_common::detect::Limits::default(),
            )?;
            ensure_path_source_current(source.as_ref(), version)?;
            match mime {
                Some(format) => {
                    let catalog = litchi_odf_common::detect::packaged_has_ooxml_catalog_read_at(
                        source.as_ref(),
                        litchi_odf_common::detect::CatalogProbeLimits::default()
                            .with_neutral_input_budget(super::super::NEUTRAL_MAX_INPUT_BYTES),
                    )?;
                    ensure_path_source_current(source.as_ref(), version)?;
                    Some((format, catalog))
                },
                None => None,
            }
        };
        #[cfg(not(any(feature = "odt", feature = "ods", feature = "odp", odf_objects)))]
        let odf: Option<(FileFormat, Option<bool>)> = None;

        if let Some((format, Some(false))) = odf {
            return odf_source(path, format, source, version, length, options);
        }

        // 2. An OOXML package: one source-backed OPC index, handed to the
        //    format's handle.
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        {
            let limits = options.ooxml_limits.unwrap_or_default();
            if length > limits.max_input_bytes() {
                return Err(litchi_opc::OpcError::Limit {
                    kind: litchi_opc::SourceLimitKind::InputBytes,
                    observed: length,
                    maximum: limits.max_input_bytes(),
                }
                .into());
            }
            #[cfg(any(test, feature = "test-support"))]
            crate::detect::record_opc_probe();
            let probe = match options.ooxml_cache {
                Some(cache) => litchi_opc::SourceBackedPackage::from_read_at(
                    Arc::clone(&source),
                    limits,
                    cache,
                    litchi_opc::SourceReadPolicy::exact(),
                    None,
                ),
                None => litchi_opc::SourceBackedPackage::from_read_at(
                    Arc::clone(&source),
                    limits,
                    litchi_opc::SourceCacheLimits::default(),
                    litchi_opc::SourceReadPolicy::exact(),
                    None,
                ),
            };
            match probe {
                Ok(package) => {
                    let format =
                        crate::detect::ooxml::try_detect_ooxml_format_from_source_backed_package(
                            &package,
                        );
                    ensure_path_source_current(source.as_ref(), version)?;
                    if let Some(format) = format? {
                        let origin = {
                            #[cfg(any(
                                feature = "docx",
                                feature = "xlsx",
                                feature = "pptx",
                                feature = "xlsb"
                            ))]
                            {
                                matches!(
                                    format,
                                    FileFormat::Docx
                                        | FileFormat::Xlsx
                                        | FileFormat::Pptx
                                        | FileFormat::Xlsb
                                )
                                .then(|| super::stable_origin(path))
                                .transpose()?
                            }
                            #[cfg(not(any(
                                feature = "docx",
                                feature = "xlsx",
                                feature = "pptx",
                                feature = "xlsb"
                            )))]
                            {
                                let _ = (format, path);
                                None
                            }
                        };
                        let file = super::super::ooxml_from_source_package_with_origin(
                            format, package, options, origin,
                        )?;
                        ensure_path_source_current(source.as_ref(), version)?;
                        return Ok(file);
                    }
                },
                Err(error) => {
                    ensure_path_source_current(source.as_ref(), version)?;
                    let odf_format = odf.map(|(format, _)| format);
                    if super::super::ooxml_format_for_extension(path).is_some()
                        || !super::super::ooxml_probe_falls_back(&error, odf_format)
                    {
                        return Err(error.into());
                    }
                },
            }
        }

        // 3. An ODF-marked package the OPC probe declined.
        if let Some((format, _)) = odf {
            return odf_source(path, format, source, version, length, options);
        }

        // 4. The lowest-precedence candidates read the input once: an ODF
        //    package without a `mimetype` member, then the tail.
        let bytes = super::super::read_whole_source(
            source.as_ref(),
            version,
            super::super::input_ceiling(options),
        )?;
        let file = super::super::bytes::zip_fallback(bytes, options, &detection_limits, tail)?;
        // An iWork package saves back to its path (record 0296).
        #[cfg(feature = "iwork-adapter")]
        if file.iwork().is_some() {
            return Ok(super::super::attach_origin(
                file,
                super::stable_origin(path)?,
            ));
        }
        Ok(file)
    }

    /// Adopt the positional owner of an ODF package of `format`.
    #[cfg_attr(
        not(any(feature = "odt", feature = "ods", feature = "odp", odf_objects)),
        allow(
            unused_variables,
            clippy::needless_pass_by_value,
            reason = "without an ODF format there is no ODF owner to build"
        )
    )]
    fn odf_source(
        path: &Path,
        format: FileFormat,
        source: Arc<dyn ReadAt>,
        version: SourceVersion,
        length: u64,
        options: &OpenOptions,
    ) -> Result<File> {
        let error = |error: litchi_core::Error| crate::Error::from(error).with_format(format);
        match format {
            #[cfg(feature = "odt")]
            FileFormat::Odt => {
                // An OOXML file name selects the OPC input ceiling even for
                // the ODT owner, as the facade's document route applied it.
                #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
                if super::super::ooxml_format_for_extension(path).is_some() {
                    let limits = options.ooxml_limits.unwrap_or_default();
                    if length > limits.max_input_bytes() {
                        return Err(litchi_opc::OpcError::Limit {
                            kind: litchi_opc::SourceLimitKind::InputBytes,
                            observed: length,
                            maximum: limits.max_input_bytes(),
                        }
                        .into());
                    }
                }
                let candidate =
                    crate::detect::detected::OdtSourcePathCandidate::new(source, version);
                Ok(File::Odt(
                    crate::format::odt::Odt::from_source_candidate(candidate, options)?
                        .with_origin(Some(super::stable_origin(path)?)),
                ))
            },
            #[cfg(feature = "ods")]
            FileFormat::Ods => {
                let spreadsheet = litchi_ods::SourceBackedSpreadsheet::from_read_at_with_options(
                    Arc::clone(&source),
                    options.odf_open_options(),
                )
                .map_err(litchi_core::Error::from)
                .map_err(error)?;
                let observed = spreadsheet
                    .source_version()
                    .map_err(crate::format::ods::odf_error)?;
                if observed != version {
                    return Err(error(litchi_core::Error::SourceChanged {
                        expected: version,
                        observed,
                    }));
                }
                Ok(File::Ods(
                    crate::format::ods::Ods::from_source(Box::new(spreadsheet), options)?
                        .with_origin(Some(super::stable_origin(path)?)),
                ))
            },
            #[cfg(feature = "odp")]
            FileFormat::Odp => {
                let presentation = litchi_odp::SourceBackedPresentation::from_read_at_with_options(
                    Arc::clone(&source),
                    options.odf_open_options(),
                )
                .map_err(litchi_core::Error::from);
                ensure_path_source_current(source.as_ref(), version)?;
                Ok(File::Odp(
                    crate::format::odp::Odp::from_source(presentation.map_err(error)?, options)?
                        .with_origin(Some(super::stable_origin(path)?)),
                ))
            },
            // The other packaged ODF formats (record 0282) have no
            // source-backed owner: their packages are read whole once and
            // adopted through the detection-prepared byte route.
            #[cfg(any(feature = "odt", odf_objects))]
            other => {
                let _ = (error, length);
                let bytes = super::super::read_whole_source(
                    source.as_ref(),
                    version,
                    super::super::input_ceiling(options),
                )?;
                let file = match litchi_odf_common::detect::prepared_or_original(
                    bytes,
                    options.odf_open_options().source_limits_value(),
                ) {
                    Ok(prepared) => super::super::odf_from_prepared(prepared, options)?,
                    Err(bytes) => {
                        litchi_odf_common::OwnedPackage::from_bytes_with_options(
                            bytes,
                            options.odf_open_options(),
                        )
                        .map_err(litchi_odf_common::Error::from)
                        .map_err(|error| crate::Error::from(error).with_format(other))?;
                        return Err(super::super::not_enabled(other));
                    },
                };
                Ok(super::super::attach_origin(
                    file,
                    super::stable_origin(path)?,
                ))
            },
            #[cfg(not(any(feature = "odt", odf_objects)))]
            other => {
                let _ = (error, path, length);
                Err(super::super::not_enabled(other))
            },
        }
    }
}
