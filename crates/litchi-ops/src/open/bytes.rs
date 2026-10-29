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
    // A valid iWork package has a canonical central-directory member that no
    // ODF or OOXML package needs.  Inspect that directory before the heavier
    // precedence probes, but only through bounded central records: no member
    // payload is read or decompressed here.  The full detector remains the
    // authority, so a false marker simply resumes the ordinary ZIP route.
    // This early route is compiled only for an all-iWork build; partial
    // feature builds retain the existing disabled-format classification.
    #[cfg(all(feature = "pages", feature = "keynote", feature = "numbers"))]
    let bytes = if likely_iwork_package(&bytes) {
        match crate::format::iwork::Iwork::detect_bytes(bytes, options)? {
            crate::format::iwork::Detected::Opened(file) => return Ok(file.into_file()),
            crate::format::iwork::Detected::Disabled(format) => {
                return Err(Error::unsupported_format(Some(format)).with_suggestion(
                    match format {
                        FileFormat::Pages => "enable the `pages` cargo feature",
                        FileFormat::Keynote => "enable the `keynote` cargo feature",
                        _ => "enable the `numbers` cargo feature",
                    },
                ));
            },
            crate::format::iwork::Detected::Declined(rejected) => rejected,
        }
    } else {
        bytes
    };
    #[cfg(not(all(feature = "pages", feature = "keynote", feature = "numbers")))]
    let bytes = bytes;

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

/// Whether the ZIP central directory contains the exact iWork marker and no
/// higher-precedence ODF/OOXML marker.  This is deliberately a marker probe,
/// not a classifier: the iWork detector still validates the package and
/// classifies its application before ownership is transferred.
#[cfg(all(feature = "pages", feature = "keynote", feature = "numbers"))]
fn likely_iwork_package(bytes: &[u8]) -> bool {
    const MAX_MARKER_DIRECTORY_BYTES: u64 = 8 * 1024 * 1024;
    const MAX_MARKER_ENTRIES: u64 = 100_000;
    const IWORK_DOCUMENT: &[u8] = b"Index/Document.iwa";
    const OOXML_CONTENT_TYPES: &[u8] = b"[Content_Types].xml";
    const ODF_MIMETYPE: &[u8] = b"mimetype";

    let Ok(archive) = soapberry_zip::ZipArchive::from_slice(bytes) else {
        return false;
    };
    if archive.entries_hint() > MAX_MARKER_ENTRIES {
        return false;
    }
    let Some(directory_bytes) = archive
        .head_eocd_offset()
        .checked_sub(archive.directory_offset())
    else {
        return false;
    };
    if directory_bytes > MAX_MARKER_DIRECTORY_BYTES {
        return false;
    }

    let mut iwork = false;
    let mut observed_entries = 0_u64;
    for entry in archive.entries() {
        observed_entries = observed_entries.saturating_add(1);
        if observed_entries > MAX_MARKER_ENTRIES {
            return false;
        }
        let Ok(entry) = entry else {
            // A malformed central record is left to the ordinary detector,
            // which owns the typed archive error and precedence behavior.
            return false;
        };
        match entry.file_path().as_bytes() {
            IWORK_DOCUMENT => iwork = true,
            OOXML_CONTENT_TYPES | ODF_MIMETYPE => return false,
            _ => {},
        }
    }
    iwork
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

#[cfg(all(test, feature = "pages", feature = "keynote", feature = "numbers"))]
mod marker_tests {
    use std::io::{Cursor, Write};

    use super::likely_iwork_package;

    fn archive(names: &[&str]) -> Vec<u8> {
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for name in names {
            writer.start_file(*name, options).expect("start ZIP member");
            writer
                .write_all(b"marker-only test payload")
                .expect("write ZIP member");
        }
        writer.finish().expect("finish ZIP").into_inner()
    }

    #[test]
    fn marker_reads_exact_central_names_only() {
        assert!(likely_iwork_package(&archive(&["Index/Document.iwa"])));
        assert!(!likely_iwork_package(&archive(&[
            "Data/Index/Document.iwa"
        ])));
        assert!(!likely_iwork_package(&archive(&["Index/Document.iwa.bak"])));
        assert!(!likely_iwork_package(&archive(&["opaque"])));
    }

    #[test]
    fn marker_does_not_bypass_higher_precedence_or_payload_decoys() {
        assert!(!likely_iwork_package(&archive(&[
            "Index/Document.iwa",
            "[Content_Types].xml",
        ])));
        assert!(!likely_iwork_package(&archive(&[
            "Index/Document.iwa",
            "mimetype"
        ])));

        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer
            .start_file("opaque", options)
            .expect("start ZIP member");
        writer
            .write_all(b"Index/Document.iwa [Content_Types].xml mimetype")
            .expect("write ZIP member");
        let bytes = writer.finish().expect("finish ZIP").into_inner();
        assert!(!likely_iwork_package(&bytes));
    }

    #[test]
    fn marker_rejects_malformed_central_records() {
        let mut bytes = archive(&["Index/Document.iwa"]);
        let central_offset = {
            let archive = soapberry_zip::ZipArchive::from_slice(&bytes).expect("locate ZIP");
            usize::try_from(archive.directory_offset()).expect("central offset")
        };
        bytes[central_offset] ^= 0xff;
        assert!(!likely_iwork_package(&bytes));
    }

    #[test]
    fn marker_counts_central_records_when_the_eocd_hint_lies() {
        const RECORDS: usize = 100_001;
        let source = archive(&["Index/Document.iwa"]);
        let source_archive = soapberry_zip::ZipArchive::from_slice(&source).expect("locate ZIP");
        let central_start = usize::try_from(source_archive.directory_offset()).expect("central");
        let central_end = usize::try_from(source_archive.head_eocd_offset()).expect("eocd");
        let record = &source[central_start..central_end];
        let mut bytes = Vec::with_capacity(
            central_start
                .saturating_add(record.len().saturating_mul(RECORDS))
                .saturating_add(22),
        );
        bytes.extend_from_slice(&source[..central_start]);
        for _ in 0..RECORDS {
            bytes.extend_from_slice(record);
        }
        let eocd_start = bytes.len();
        bytes.extend_from_slice(&source[central_end..]);
        let central_size = u32::try_from(record.len().saturating_mul(RECORDS)).expect("size");
        let central_start = u32::try_from(central_start).expect("offset");
        bytes[eocd_start + 8..eocd_start + 10].copy_from_slice(&0_u16.to_le_bytes());
        bytes[eocd_start + 10..eocd_start + 12].copy_from_slice(&0_u16.to_le_bytes());
        bytes[eocd_start + 12..eocd_start + 16].copy_from_slice(&central_size.to_le_bytes());
        bytes[eocd_start + 16..eocd_start + 20].copy_from_slice(&central_start.to_le_bytes());
        assert!(!likely_iwork_package(&bytes));
    }

    #[test]
    fn marker_probe_respects_the_outer_input_ceiling() {
        let source = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/pages/body-table-comment-hidden-native-saved.pages"),
        )
        .expect("read tracked iWork fixture");
        let error = crate::open_bytes_with(source, &crate::OpenOptions::new().max_input_bytes(1))
            .expect_err("tight input ceiling must refuse before marker routing");
        assert_eq!(error.code(), crate::Code::LimitExceeded);
    }
}

#[cfg(all(
    test,
    feature = "pages",
    not(feature = "keynote"),
    not(feature = "numbers")
))]
mod disabled_marker_tests {
    #[test]
    fn partial_iwork_build_keeps_disabled_format_refusal() {
        let source = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/keynote/basic.key"),
        )
        .expect("read tracked Keynote fixture");
        let error = crate::open_bytes(source).expect_err("Keynote is disabled in this build");
        assert_eq!(error.code(), crate::Code::UnsupportedFormat);
        assert_eq!(error.detected(), Some(litchi_core::FileFormat::Keynote));
    }
}
