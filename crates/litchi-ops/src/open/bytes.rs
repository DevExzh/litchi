//! The in-memory route of [`crate::open_bytes`], and the byte fallback the
//! path route takes for inputs it reads whole.

use litchi_core::FileFormat;
use litchi_core::detection::simd_utils::check_office_signatures;

use crate::detect::tail::ZipTail;
use crate::error::{Error, Result};
use crate::{File, OpenOptions};

/// Classify `bytes` and adopt the owner its format keeps. `ooxml_hint` is
/// the OOXML format a file name named, for an encrypted package.
#[cfg_attr(
    not(any(
        feature = "doc",
        feature = "ppt",
        feature = "xls",
        feature = "encryption"
    )),
    allow(
        unused_variables,
        reason = "only an encrypted OOXML package in a compound file uses the hint"
    )
)]
pub(super) fn open<T: ZipTail>(
    bytes: Vec<u8>,
    options: &OpenOptions,
    tail: &T,
    ooxml_hint: Option<FileFormat>,
) -> Result<File> {
    let detection_limits = options.detection_limits();
    if let Some(limit) = options.max_input_bytes {
        super::check_input_len(u64::try_from(bytes.len()).unwrap_or(u64::MAX), limit)?;
    }
    if bytes.len() < 4 {
        return Err(Error::unsupported_format(None));
    }

    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    if let Some(format) =
        litchi_odf_common::detect::flat(&bytes, litchi_odf_common::detect::Limits::default())
            .ok()
            .flatten()
    {
        let maximum = detection_limits.odf().max_input_bytes();
        let observed = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
        if observed > maximum {
            return Err(Error::from(litchi_odf_common::Error::Limit {
                kind: litchi_odf_common::LimitKind::Source(
                    litchi_odf_common::SourceLimitKind::InputBytes,
                ),
                observed,
                maximum,
            })
            .with_format(format));
        }
        #[cfg(feature = "ods")]
        if format == FileFormat::Ods {
            return Ok(File::Ods(crate::format::ods::Ods::from_flat(
                bytes, options,
            )?));
        }
        #[cfg(feature = "odp")]
        if format == FileFormat::Odp {
            return Ok(File::Odp(crate::format::odp::Odp::from_flat(
                bytes, options,
            )?));
        }
        #[cfg(feature = "odt")]
        if format == FileFormat::Odt {
            return Ok(File::Odt(crate::format::odt::Odt::from_flat(
                bytes, options,
            )?));
        }
        #[cfg(feature = "odg")]
        if format == FileFormat::Odg {
            return Ok(File::Odg(crate::format::odf_tree::OdfTree::from_flat(
                bytes, options,
            )?));
        }
        return Err(super::flat_odf_unsupported(format));
    }

    let mask = check_office_signatures(&bytes);
    #[cfg(feature = "rtf")]
    if mask.is_rtf() {
        return Ok(File::Rtf(crate::format::rtf::Rtf::from_bytes(
            bytes, options,
        )?));
    }
    if mask.is_ole2() {
        #[cfg(any(
            feature = "doc",
            feature = "ppt",
            feature = "xls",
            feature = "encryption"
        ))]
        return super::cfb::open_bytes(bytes, options, ooxml_hint);
        #[cfg(not(any(
            feature = "doc",
            feature = "ppt",
            feature = "xls",
            feature = "encryption"
        )))]
        return Err(Error::unsupported_format(None));
    }
    // Compressed RTF comes after the container formats, so the ZIP and OLE
    // precedence is unchanged for crafted polyglots.
    #[cfg(feature = "rtf")]
    if !mask.is_zip() && crate::detect::is_compressed_rtf(&bytes) {
        return Ok(File::Rtf(crate::format::rtf::Rtf::from_bytes(
            bytes, options,
        )?));
    }
    if mask.is_zip() {
        return zip(bytes, options, &detection_limits, tail);
    }
    Err(Error::unsupported_format(None))
}

/// Classify an in-memory ZIP: an ordinary ODF package first, then OOXML
/// through one OPC index, then the lower-precedence candidates.
fn zip<T: ZipTail>(
    bytes: Vec<u8>,
    options: &OpenOptions,
    detection_limits: &crate::detect::SourceLimits,
    tail: &T,
) -> Result<File> {
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    let odf = litchi_odf_common::detect::packaged_mime(
        &bytes,
        litchi_odf_common::detect::Limits::default(),
    )
    .ok()
    .flatten();
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    if odf.is_some()
        && litchi_odf_common::detect::packaged_has_ooxml_catalog(
            &bytes,
            litchi_odf_common::detect::CatalogProbeLimits::default()
                .with_neutral_input_budget(super::NEUTRAL_MAX_INPUT_BYTES),
        ) == Some(false)
    {
        return zip_fallback(bytes, options, detection_limits, tail);
    }
    #[cfg(not(any(feature = "odt", feature = "ods", feature = "odp", odf_objects)))]
    let odf: Option<FileFormat> = None;

    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    let bytes = match ooxml(bytes, options, detection_limits, odf)? {
        Ok(file) => return Ok(file),
        Err(bytes) => bytes,
    };
    #[cfg(not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")))]
    let _ = odf;
    zip_fallback(bytes, options, detection_limits, tail)
}

/// Probe an in-memory ZIP as OOXML through exactly one OPC index. A package
/// that is not OOXML gives its bytes back for the lower-precedence
/// candidates.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn ooxml(
    bytes: Vec<u8>,
    options: &OpenOptions,
    detection_limits: &crate::detect::SourceLimits,
    odf: Option<FileFormat>,
) -> Result<std::result::Result<File, Vec<u8>>> {
    use std::sync::Arc;

    let limits = detection_limits.ooxml();
    if options.source_mode == crate::SourceMode::Eager {
        // The eager detection route retains the caller's owned archive.  A
        // borrowed `from_bytes` package has no preservation records, so an
        // eager no-op would silently canonicalize the ZIP on publication.
        #[cfg(any(test, feature = "test-support"))]
        crate::detect::record_opc_probe();
        let shared = Arc::new(bytes);
        return match litchi_opc::OpcPackage::from_shared_vec(Arc::clone(&shared), limits) {
            Ok(package) => match crate::detect::ooxml::detect_ooxml_format_from_package(&package) {
                Some(format) => Ok(Ok(super::ooxml_from_eager_package(
                    format, package, options,
                )?)),
                None => {
                    drop(package);
                    Ok(Err(reclaim(shared)))
                },
            },
            Err(error) if super::ooxml_probe_falls_back(&error, odf) => Ok(Err(reclaim(shared))),
            Err(error) => Err(error.into()),
        };
    }

    // Keep a reclaimable owner, so a ZIP that is not OOXML continues to the
    // lower-precedence candidates without a copy.
    let shared = Arc::new(bytes);
    let source: Arc<dyn litchi_core::ReadAt> =
        Arc::new(litchi_core::OwnedSource::from_arc(Arc::clone(&shared)));
    #[cfg(any(test, feature = "test-support"))]
    crate::detect::record_opc_probe();
    let probe = match options.ooxml_cache {
        Some(cache) => litchi_opc::SourceBackedPackage::from_read_at(
            source,
            limits,
            cache,
            litchi_opc::SourceReadPolicy::exact(),
            None,
        ),
        None => litchi_opc::SourceBackedPackage::from_read_at(
            source,
            limits,
            litchi_opc::SourceCacheLimits::default(),
            litchi_opc::SourceReadPolicy::exact(),
            None,
        ),
    };
    match probe {
        Ok(package) => {
            match crate::detect::ooxml::try_detect_ooxml_format_from_source_backed_package(
                &package,
            )? {
                Some(format) => {
                    drop(shared);
                    Ok(Ok(super::ooxml_from_source_package(
                        format, package, options,
                    )?))
                },
                None => {
                    drop(package);
                    Ok(Err(reclaim(shared)))
                },
            }
        },
        Err(error) if super::ooxml_probe_falls_back(&error, odf) => Ok(Err(reclaim(shared))),
        Err(error) => Err(error.into()),
    }
}

/// Take the input back once every source over it has been dropped. A copy
/// is the fallback only if a leaf still held a reference.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn reclaim(shared: std::sync::Arc<Vec<u8>>) -> Vec<u8> {
    std::sync::Arc::try_unwrap(shared).unwrap_or_else(|shared| shared.as_ref().clone())
}

/// The lowest-precedence ZIP candidates: an ODF package (with or without a
/// `mimetype` member), then an Apple iWork package (in a build with an iWork
/// feature), then the caller's tail.
pub(super) fn zip_fallback<T: ZipTail>(
    bytes: Vec<u8>,
    options: &OpenOptions,
    detection_limits: &crate::detect::SourceLimits,
    tail: &T,
) -> Result<File> {
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    let bytes = match litchi_odf_common::detect::prepared_or_original(bytes, detection_limits.odf())
    {
        Ok(prepared) => return super::odf_from_prepared(prepared, options),
        Err(bytes) => {
            // Preparation returns ownership on rejection. Once the bounded
            // MIME probe identifies ODF, let its ordinary owner report the
            // typed source failure instead of turning it into an unknown
            // format. Successful preparation still transfers its one index.
            match litchi_odf_common::detect::packaged_mime(
                &bytes,
                litchi_odf_common::detect::Limits::default(),
            )
            .ok()
            .flatten()
            {
                Some(format) => litchi_odf_common::OwnedPackage::from_bytes_with_options(
                    bytes,
                    litchi_odf_common::OpenOptions::new().source_limits(detection_limits.odf()),
                )
                .map_err(litchi_odf_common::Error::from)
                .map_err(|error| Error::from(error).with_format(format))?
                .into_inner(),
                None => bytes,
            }
        },
    };
    #[cfg(not(any(feature = "odt", feature = "ods", feature = "odp", odf_objects)))]
    let _ = (options, detection_limits);
    #[cfg(feature = "iwork-adapter")]
    let bytes = match crate::format::iwork::Iwork::detect_bytes(bytes, options)? {
        crate::format::iwork::Detected::Opened(file) => return Ok(file.into_file()),
        crate::format::iwork::Detected::Disabled(format) => {
            return Err(
                Error::unsupported_format(Some(format)).with_suggestion(match format {
                    FileFormat::Pages => "enable the `pages` cargo feature",
                    FileFormat::Keynote => "enable the `keynote` cargo feature",
                    _ => "enable the `numbers` cargo feature",
                }),
            );
        },
        crate::format::iwork::Detected::Declined(bytes) => bytes,
    };
    match tail.classify_bytes(&bytes) {
        Some(format) => Err(Error::iwork_unsupported(format)),
        None => Err(Error::unsupported_format(None)),
    }
}
