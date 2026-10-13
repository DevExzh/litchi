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
    if let Some(limit) = options.max_input_bytes {
        super::check_input_len(u64::try_from(bytes.len()).unwrap_or(u64::MAX), limit)?;
    }
    if bytes.len() < 4 {
        return Err(Error::unsupported_format(None));
    }

    #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
    if let Some(format) = litchi_odf_common::detect::flat(&bytes) {
        #[cfg(feature = "ods")]
        if format == FileFormat::Ods {
            return Ok(File::Ods(crate::format::ods::Ods::from_flat(
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
        return zip(bytes, options, tail);
    }
    Err(Error::unsupported_format(None))
}

/// Classify an in-memory ZIP: an ordinary ODF package first, then OOXML
/// through one OPC index, then the lower-precedence candidates.
fn zip<T: ZipTail>(bytes: Vec<u8>, options: &OpenOptions, tail: &T) -> Result<File> {
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
    let odf = litchi_odf_common::detect::packaged_mime(&bytes);
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
    if odf.is_some() && litchi_odf_common::detect::packaged_has_ooxml_catalog(&bytes) == Some(false)
    {
        return zip_fallback(bytes, options, tail);
    }
    #[cfg(not(any(feature = "odt", feature = "ods", feature = "odp")))]
    let odf: Option<FileFormat> = None;

    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    let bytes = match ooxml(bytes, options, odf)? {
        Ok(file) => return Ok(file),
        Err(bytes) => bytes,
    };
    #[cfg(not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")))]
    let _ = odf;
    zip_fallback(bytes, options, tail)
}

/// Probe an in-memory ZIP as OOXML through exactly one OPC index. A package
/// that is not OOXML gives its bytes back for the lower-precedence
/// candidates.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn ooxml(
    bytes: Vec<u8>,
    options: &OpenOptions,
    odf: Option<FileFormat>,
) -> Result<std::result::Result<File, Vec<u8>>> {
    use std::sync::Arc;

    let limits = options.ooxml_limits.unwrap_or_default();
    if options.source_mode == crate::SourceMode::Eager {
        // The eager detection route retains the caller's owned archive.  A
        // borrowed `from_bytes` package has no preservation records, so an
        // eager no-op would silently canonicalize the ZIP on publication.
        #[cfg(any(test, feature = "test-support"))]
        crate::detect::record_opc_probe();
        let shared = Arc::new(bytes);
        return match litchi_opc::OpcPackage::from_shared_vec_with_limits(
            Arc::clone(&shared),
            limits,
        ) {
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
        Some(cache) => litchi_opc::SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
            source, limits, cache,
        ),
        None => litchi_opc::SourceBackedPackage::from_read_at_with_limits(source, limits),
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
/// `mimetype` member), then the caller's tail, which names iWork packages.
pub(super) fn zip_fallback<T: ZipTail>(
    bytes: Vec<u8>,
    options: &OpenOptions,
    tail: &T,
) -> Result<File> {
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
    let bytes = match litchi_odf_common::detect::prepared_or_original(bytes) {
        Ok(prepared) => return super::odf_from_prepared(prepared, options),
        Err(bytes) => bytes,
    };
    #[cfg(not(any(feature = "odt", feature = "ods", feature = "odp")))]
    let _ = options;
    match tail.classify_bytes(&bytes) {
        Some(format) => Err(Error::iwork_unsupported(format)),
        None => Err(Error::unsupported_format(None)),
    }
}
