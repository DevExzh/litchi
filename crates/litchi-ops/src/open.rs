//! The single-pass openers: [`open`], [`open_bytes`], [`open_prepared`] and
//! their `_with` variants (record 0003 §2.5 and §2.7).
//!
//! Each call runs one detection pass and hands what detection built (the
//! positional source, the OPC package, the compound file, the ODF package
//! index) to the matching handle constructor, so an input is indexed once.
//! Precedence is the facade's: a compound file is classified by its host
//! stream; a ZIP is an ordinary ODF package when its `mimetype` names one and
//! it carries no OPC catalog, an OOXML package when its OPC catalog names one,
//! then any other ODF package, and last an iWork package, which only the
//! `litchi` facade's tail can name; RTF is recognized by its header.

use std::path::Path;
use std::sync::Arc;

use litchi_core::FileFormat;

use crate::detect::tail::{NoZipTail, ZipTail};
use crate::error::{Error, Result};
use crate::{File, OpenOptions};

mod bytes;
#[cfg(any(
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "encryption"
))]
mod cfb;
mod path;

/// The input ceiling every route applies when [`OpenOptions::max_input_bytes`]
/// is unset: the facade's `UNIFIED_*_FALLBACK_MAX_INPUT_BYTES`.
const NEUTRAL_MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;

/// Open the file at `path`, detecting its format from the content.
///
/// One detection pass classifies the file and builds the owner the format's
/// handle keeps: on `unix` and `windows` most formats stay source-backed,
/// reading payloads from the file on demand; legacy Word files, RTF and the
/// in-memory routes read the file once. An encrypted input opens as
/// [`File::Locked`] and a signed one as [`File::Signed`].
///
/// # Errors
///
/// Returns [`Code::FileNotFound`](crate::Code::FileNotFound) when the path
/// does not exist,
/// [`Code::UnsupportedFormat`](crate::Code::UnsupportedFormat) with no
/// [`detected`](Error::detected) format for an input that is not an Office
/// file, or with the detected format when its cargo feature is off (an
/// Apple iWork package or package directory included; a package directory
/// opens as its single-file form and `save()` writes it back as a
/// directory, record 0306), [`Code::LimitExceeded`](crate::Code::LimitExceeded)
/// when the input exceeds a read limit, and the format's typed error when its
/// container cannot be read. In a build with an OOXML format, a file named
/// as an OOXML package (`.docx`, `.xlsx`, `.pptx`, `.xlsb` and their macro
/// and template forms) whose content is neither a ZIP package nor an
/// encrypted OOXML package is refused with
/// [`Code::CorruptFile`](crate::Code::CorruptFile), as the facade refused it.
///
/// ```no_run
/// # fn main() -> litchi_ops::Result<()> {
/// let file = litchi_ops::open("report.docx")?;
/// println!("{}", file.text()?);
/// # Ok(())
/// # }
/// ```
pub fn open(path: impl AsRef<Path>) -> Result<File> {
    open_with(path, &OpenOptions::default())
}

/// [`open`] with explicit [`OpenOptions`]: input and per-format limits,
/// owner selection, source retention and an execution context.
///
/// # Errors
///
/// As [`open`], plus [`Code::Cancelled`](crate::Code::Cancelled) when the
/// execution context is cancelled before or after the open.
///
/// ```no_run
/// use litchi_ops::{OpenOptions, SourceMode};
///
/// # fn main() -> litchi_ops::Result<()> {
/// let options = OpenOptions::new()
///     .max_input_bytes(64 << 20)
///     .source_mode(SourceMode::Eager);
/// let file = litchi_ops::open_with("report.docx", &options)?;
/// # let _ = file;
/// # Ok(())
/// # }
/// ```
pub fn open_with(path: impl AsRef<Path>, options: &OpenOptions) -> Result<File> {
    open_with_tail(path, options, &NoZipTail)
}

/// Open an in-memory input, detecting its format from the content.
///
/// The bytes move into the handle without a copy: source-backed formats
/// read them in place, and a legacy binary document keeps them as its edit
/// source. Pass a `Vec<u8>` you own to avoid a copy; a borrowed slice is
/// copied once.
///
/// # Errors
///
/// As [`open`], without the file-system errors.
///
/// ```
/// # #[cfg(feature = "rtf")]
/// # {
/// let file = litchi_ops::open_bytes(br"{\rtf1 Hello}".to_vec()).unwrap();
/// assert_eq!(file.text().unwrap().trim(), "Hello");
/// # }
/// ```
pub fn open_bytes(bytes: impl Into<Vec<u8>>) -> Result<File> {
    open_bytes_with(bytes, &OpenOptions::default())
}

/// [`open_bytes`] with explicit [`OpenOptions`].
///
/// # Errors
///
/// As [`open_bytes`], plus [`Code::Cancelled`](crate::Code::Cancelled) when
/// the execution context is cancelled before or after the open.
///
/// ```
/// # #[cfg(feature = "rtf")]
/// # {
/// use litchi_ops::{Code, OpenOptions};
///
/// let options = OpenOptions::new().max_input_bytes(4);
/// let error = litchi_ops::open_bytes_with(br"{\rtf1 Hello}".to_vec(), &options).unwrap_err();
/// assert_eq!(error.code(), Code::LimitExceeded);
/// # }
/// ```
pub fn open_bytes_with(bytes: impl Into<Vec<u8>>, options: &OpenOptions) -> Result<File> {
    open_bytes_with_tail(bytes, options, &NoZipTail)
}

/// [`open_with`] with a lowest-precedence ZIP classifier.
///
/// A ZIP input that no other candidate claims (neither an OOXML nor an
/// OpenDocument package) is handed to `tail`; see
/// [`detect::tail`](crate::detect::tail). An input the tail names is refused
/// with [`Error::iwork_unsupported`], with the format the tail reported.
/// Apple iWork packages no longer need a tail: in a build with the `pages`,
/// `keynote` or `numbers` feature, detection classifies them itself, as the
/// last built-in ZIP candidate before the tail (record 0296).
///
/// # Errors
///
/// As [`open_with`], plus [`Code::UnsupportedFormat`](crate::Code::UnsupportedFormat)
/// with the tail's format for an input the tail names.
///
/// ```no_run
/// use litchi_ops::{OpenOptions, detect::SourceLimits};
/// use litchi_ops::detect::tail::NoZipTail;
///
/// # fn main() -> litchi_ops::Result<()> {
/// // With the empty tail this is exactly `open_with`.
/// let file = litchi_ops::open_with_tail("report.docx", &OpenOptions::new(), &NoZipTail)?;
/// # let _ = file;
/// # Ok(())
/// # }
/// ```
pub fn open_with_tail<T: ZipTail>(
    path: impl AsRef<Path>,
    options: &OpenOptions,
    tail: &T,
) -> Result<File> {
    check_cancelled(options)?;
    let file = match path::open(path.as_ref(), options, tail) {
        Ok(file) => file,
        Err(error) => open_directory(path.as_ref(), options, error)?,
    };
    check_cancelled(options)?;
    Ok(file)
}

/// Open the iWork package directory at `path` the file route refused: the
/// file route reads one positional source, so an app-authored package
/// directory fails there and is classified here, on the error path only. A
/// Pages, Keynote or Numbers directory opens through its leaf (records 0299
/// and 0306); a directory of a format whose feature is off is
/// `unsupported_format` naming the feature.
#[cfg_attr(
    not(feature = "iwork-adapter"),
    allow(
        clippy::unnecessary_wraps,
        reason = "the directory route exists only with an iWork format"
    )
)]
fn open_directory(path: &Path, options: &OpenOptions, error: Error) -> Result<File> {
    #[cfg(feature = "iwork-adapter")]
    if path.is_dir()
        && let Some(handle) = crate::format::iwork::Iwork::capture_directory(path, options)?
    {
        return Ok(handle.into_file());
    }
    #[cfg(feature = "iwork-adapter")]
    if path.is_dir()
        && let Some(format) = crate::format::iwork::classify_directory(path)
    {
        return match format {
            #[cfg(feature = "pages")]
            FileFormat::Pages => crate::format::iwork::Iwork::open_directory(path, format, options)
                .map(crate::format::iwork::Iwork::into_file),
            #[cfg(feature = "keynote")]
            FileFormat::Keynote => {
                crate::format::iwork::Iwork::open_directory(path, format, options)
                    .map(crate::format::iwork::Iwork::into_file)
            },
            #[cfg(feature = "numbers")]
            FileFormat::Numbers => {
                crate::format::iwork::Iwork::open_directory(path, format, options)
                    .map(crate::format::iwork::Iwork::into_file)
            },
            _ => Err(
                Error::unsupported_format(Some(format)).with_suggestion(match format {
                    FileFormat::Pages => "enable the `pages` cargo feature",
                    FileFormat::Keynote => "enable the `keynote` cargo feature",
                    _ => "enable the `numbers` cargo feature",
                }),
            ),
        };
    }
    #[cfg(not(feature = "iwork-adapter"))]
    let _ = (path, options);
    Err(error)
}

/// [`open_bytes_with`] with a lowest-precedence ZIP classifier; see
/// [`open_with_tail`].
///
/// # Errors
///
/// As [`open_bytes_with`], plus
/// [`Code::UnsupportedFormat`](crate::Code::UnsupportedFormat) with the
/// tail's format for an input the tail names.
pub fn open_bytes_with_tail<T: ZipTail>(
    bytes: impl Into<Vec<u8>>,
    options: &OpenOptions,
    tail: &T,
) -> Result<File> {
    check_cancelled(options)?;
    let file = bytes::open(bytes.into(), options, tail, None)?;
    check_cancelled(options)?;
    Ok(file)
}

/// Open the OOXML package a [`detect_and_prepare`](crate::detect::detect_and_prepare)
/// call prepared, without indexing the container again.
///
/// # Errors
///
/// Returns [`Code::InvalidArgument`](crate::Code::InvalidArgument) when the
/// detection prepared no package (it classified a non-OOXML input; open its
/// path instead), [`Code::IoError`](crate::Code::IoError) when the file
/// changed after it was prepared, and the format's typed error otherwise.
///
/// ```no_run
/// use litchi_ops::OpenOptions;
/// use litchi_ops::detect::{detect_and_prepare, SourceLimits};
///
/// # fn main() -> litchi_ops::Result<()> {
/// if let Some(detection) = detect_and_prepare("report.docx", &SourceLimits::default()) {
///     if detection.is_prepared() {
///         let file = litchi_ops::open_prepared(detection)?;
///         println!("{:?}", file.format());
///     }
/// }
/// # Ok(())
/// # }
/// ```
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub fn open_prepared(detection: crate::detect::PreparedDetection) -> Result<File> {
    open_prepared_with(detection, &OpenOptions::default())
}

/// [`open_prepared`] with explicit [`OpenOptions`]. The read limits were
/// applied when the package was prepared; owner selection and source
/// retention apply here. Under [`SourceMode::Eager`](crate::SourceMode::Eager)
/// the already indexed package is materialized into the owning form.
///
/// # Errors
///
/// As [`open_prepared`], plus [`Code::Cancelled`](crate::Code::Cancelled)
/// when the execution context is cancelled before or after the open.
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub fn open_prepared_with(
    detection: crate::detect::PreparedDetection,
    options: &OpenOptions,
) -> Result<File> {
    check_cancelled(options)?;
    let origin = detection.origin();
    let (format, prepared) = detection.into_parts();
    let Some(prepared) = prepared else {
        return Err(Error::new(
            crate::Code::InvalidArgument,
            format!("the {format:?} detection prepared no OOXML package to adopt"),
        )
        .with_format(format)
        .with_suggestion("open the path with `litchi_ops::open`"));
    };
    let package = prepared.into_package()?;
    let file = ooxml_from_source_package_with_origin(format, package, options, origin)?;
    check_cancelled(options)?;
    Ok(file)
}

/// The input ceiling: [`OpenOptions::max_input_bytes`], or the neutral
/// ceiling every facade route applied.
pub(crate) fn input_ceiling(options: &OpenOptions) -> u64 {
    options.max_input_bytes.unwrap_or(NEUTRAL_MAX_INPUT_BYTES)
}

/// Refuse an input longer than `limit`.
pub(crate) fn check_input_len(observed: u64, limit: u64) -> Result<()> {
    if observed > limit {
        return Err(
            litchi_core::Error::ResourceLimit(litchi_core::ResourceLimit {
                resource: litchi_core::Resource::InputBytes,
                observed,
                limit,
                scope: Arc::from("unified open input"),
            })
            .into(),
        );
    }
    Ok(())
}

/// Honor the execution context's cancellation at an open boundary.
fn check_cancelled(options: &OpenOptions) -> Result<()> {
    if let Some(execution) = &options.execution {
        execution.check()?;
    }
    Ok(())
}

/// Read a whole positional source under `ceiling`, checked against the
/// revision `version` before and after the read.
#[cfg(any(unix, windows))]
pub(crate) fn read_whole_source(
    source: &dyn litchi_core::ReadAt,
    version: litchi_core::SourceVersion,
    ceiling: u64,
) -> Result<Vec<u8>> {
    Ok(crate::detect::detected::read_path_source_bytes(
        source, version, ceiling,
    )?)
}

/// The OOXML format an OOXML file name extension names.
fn ooxml_format_for_extension(path: &Path) -> Option<FileFormat> {
    let extension = path.extension()?.to_str()?;
    let names = |known: &[&str]| {
        known
            .iter()
            .any(|name| extension.eq_ignore_ascii_case(name))
    };
    if names(&["docx", "docm", "dotx", "dotm"]) {
        Some(FileFormat::Docx)
    } else if names(&["pptx", "pptm", "ppsx", "ppsm", "potx", "potm"]) {
        Some(FileFormat::Pptx)
    } else if names(&["xlsx", "xlsm", "xltx", "xltm"]) {
        Some(FileFormat::Xlsx)
    } else if names(&["xlsb"]) {
        Some(FileFormat::Xlsb)
    } else {
        None
    }
}

/// Refuse an input an OOXML file name calls a package it is not: neither a
/// ZIP nor, in a compound file, an encrypted OOXML package. The facade's
/// suffix preflight refused these before any content routing.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn ooxml_suffix_mismatch() -> Error {
    litchi_opc::OpcError::ZipError("OOXML-suffixed input does not have ZIP magic".to_owned()).into()
}

/// Refuse a recognized format this build cannot open: with a suggestion to
/// enable its cargo feature when it has one.
fn not_enabled(format: FileFormat) -> Error {
    let feature = match format {
        FileFormat::Doc => "doc",
        FileFormat::Docx => "docx",
        FileFormat::Ppt => "ppt",
        FileFormat::Pptx => "pptx",
        FileFormat::Xls => "xls",
        FileFormat::Xlsx => "xlsx",
        FileFormat::Xlsb => "xlsb",
        FileFormat::Rtf => "rtf",
        FileFormat::Odt => "odt",
        FileFormat::Ods => "ods",
        FileFormat::Odp => "odp",
        FileFormat::Oth => "oth",
        FileFormat::Odm => "odm",
        FileFormat::Odg => "odg",
        FileFormat::Odc => "odc",
        FileFormat::Odi => "odi",
        FileFormat::Odb => "odb",
        FileFormat::Odf => "odf-formula",
        _ => return Error::unsupported_format(Some(format)),
    };
    Error::feature_disabled(format, feature)
}

/// Wrap `file` in [`File::Signed`] when its container carries a signature.
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb"
))]
fn signed_if(file: File, format: FileFormat, signed: bool, origin: crate::SignatureOrigin) -> File {
    if signed {
        File::Signed(crate::Signed::new(file, format, origin))
    } else {
        file
    }
}

/// Attach a path origin to a publishable owner returned by a byte fallback. The eager path
/// route reaches the byte classifier before it knows the package's format;
/// this helper applies metadata only after classification and preserves a
/// signed wrapper by temporarily using its existing explicit disposition.
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
pub(crate) fn attach_origin(file: File, origin: Arc<Path>) -> File {
    match file {
        #[cfg(feature = "rtf")]
        File::Rtf(rtf) => File::Rtf(rtf.with_origin(Some(origin))),
        #[cfg(feature = "odt")]
        File::Odt(odt) | File::Oth(odt) | File::Odm(odt) => {
            odt.with_origin(Some(origin)).into_file()
        },
        #[cfg(odf_objects)]
        File::Odg(tree) | File::Odc(tree) | File::Odi(tree) | File::Odb(tree) | File::Odf(tree) => {
            tree.with_origin(Some(origin)).into_file()
        },
        #[cfg(feature = "iwork-adapter")]
        File::Pages(iwork) | File::Keynote(iwork) | File::Numbers(iwork) => {
            iwork.with_origin(Some(origin)).into_file()
        },
        #[cfg(feature = "odp")]
        File::Odp(odp) => File::Odp(odp.with_origin(Some(origin))),
        #[cfg(feature = "doc")]
        File::Doc(doc) => File::Doc(doc.with_origin(Some(origin))),
        #[cfg(feature = "xls")]
        File::Xls(xls) => File::Xls(xls.with_origin(Some(origin))),
        #[cfg(feature = "docx")]
        File::Docx(docx) => File::Docx(docx.with_origin(Some(origin))),
        #[cfg(feature = "ppt")]
        File::Ppt(ppt) => File::Ppt(ppt.with_origin(Some(origin))),
        #[cfg(feature = "xlsx")]
        File::Xlsx(xlsx) => File::Xlsx(xlsx.with_origin(Some(origin))),
        #[cfg(feature = "xlsb")]
        File::Xlsb(xlsb) => File::Xlsb(xlsb.with_origin(Some(origin))),
        #[cfg(feature = "pptx")]
        File::Pptx(pptx) => File::Pptx(pptx.with_origin(Some(origin))),
        #[cfg(feature = "ods")]
        File::Ods(ods) => File::Ods(ods.with_origin(Some(origin))),
        #[cfg(any(
            feature = "doc",
            feature = "docx",
            feature = "ppt",
            feature = "xlsx",
            feature = "pptx",
            feature = "xls",
            feature = "xlsb"
        ))]
        File::Signed(signed) => {
            let format = signed.format();
            let signature_origin = signed.origin();
            let file = attach_origin(signed.unsign(), origin);
            File::Signed(crate::Signed::new(file, format, signature_origin))
        },
        #[cfg(any(
            feature = "doc",
            feature = "docx",
            feature = "ppt",
            feature = "xlsx",
            feature = "pptx",
            feature = "xls",
            feature = "xlsb"
        ))]
        File::Locked(locked) => File::Locked(locked.with_origin(Some(origin))),
        #[allow(unreachable_patterns)]
        other => other,
    }
}

/// Whether a failed OPC probe leaves the ZIP to the lower-precedence
/// candidates: the package has no OPC content-types catalog. An
/// ODT-marked package falls back only when the catalog part is missing
/// outright; one that carries a malformed catalog is refused with the OPC
/// error, as the facade's document route refused it.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn ooxml_probe_falls_back(error: &litchi_opc::OpcError, odf: Option<FileFormat>) -> bool {
    use litchi_opc::OpcError;

    if crate::detect::detected::hard_ooxml_probe_error(error) {
        return false;
    }
    let missing_part = matches!(
        error,
        OpcError::PartNotFound(part) if part == "[Content_Types].xml" || part == "/[Content_Types].xml"
    );
    if odf == Some(FileFormat::Odt) {
        missing_part
    } else {
        missing_part || matches!(error, OpcError::InvalidContentTypesManifest(_))
    }
}

/// Build the handle for a source-backed OOXML package of `format`.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn ooxml_from_source_package(
    format: FileFormat,
    package: litchi_opc::SourceBackedPackage,
    options: &OpenOptions,
) -> Result<File> {
    ooxml_from_source_package_with_origin(format, package, options, None)
}

/// Build a source-backed OOXML handle while retaining the path origin which
/// produced it. The origin is metadata only; package classification still
/// comes from the already-indexed package.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub(crate) fn ooxml_from_source_package_with_origin(
    format: FileFormat,
    package: litchi_opc::SourceBackedPackage,
    options: &OpenOptions,
    origin: Option<Arc<Path>>,
) -> Result<File> {
    #[cfg(not(any(feature = "docx", feature = "xlsx", feature = "pptx")))]
    let _ = &origin;
    let signed = package.has_signature_infrastructure();
    let file = match format {
        #[cfg(feature = "docx")]
        FileFormat::Docx => File::Docx(crate::format::docx::Docx::from_source_with_origin(
            litchi_docx::source_backed::Package::from_source_backed_package(package)?,
            options,
            origin,
        )?),
        #[cfg(feature = "pptx")]
        FileFormat::Pptx => File::Pptx(
            crate::format::pptx::Pptx::from_source_package(package, options)?.with_origin(origin),
        ),
        #[cfg(feature = "xlsx")]
        FileFormat::Xlsx => File::Xlsx(
            crate::format::xlsx::Xlsx::from_source_package(package, options)?.with_origin(origin),
        ),
        #[cfg(feature = "xlsb")]
        FileFormat::Xlsb => File::Xlsb(
            crate::format::xlsb::Xlsb::from_source_package(package, options)?.with_origin(origin),
        ),
        other => return Err(not_enabled(other)),
    };
    Ok(signed_if(file, format, signed, crate::SignatureOrigin::Opc))
}

/// Build the eager handle for an OOXML package of `format` the eager probe
/// parsed.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn ooxml_from_eager_package(
    format: FileFormat,
    package: litchi_opc::OpcPackage,
    options: &OpenOptions,
) -> Result<File> {
    ooxml_from_eager_package_with_origin(format, package, options, None)
}

/// Build an eager OOXML handle while retaining the path origin which
/// produced it. DOCX and XLSX store this value; the other format
/// handles keep their existing state until their save wave lands.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub(crate) fn ooxml_from_eager_package_with_origin(
    format: FileFormat,
    package: litchi_opc::OpcPackage,
    options: &OpenOptions,
    origin: Option<Arc<Path>>,
) -> Result<File> {
    #[cfg(not(any(feature = "docx", feature = "xlsx", feature = "pptx")))]
    let _ = &origin;
    let signed = package.is_signed();
    let file = match format {
        #[cfg(feature = "docx")]
        FileFormat::Docx => File::Docx(crate::format::docx::Docx::from_opc_package_with_origin(
            package, options, origin,
        )?),
        #[cfg(feature = "pptx")]
        FileFormat::Pptx => File::Pptx(
            crate::format::pptx::Pptx::from_opc_package(package, options)?.with_origin(origin),
        ),
        #[cfg(feature = "xlsx")]
        FileFormat::Xlsx => File::Xlsx(
            crate::format::xlsx::Xlsx::from_opc_package(package, options)?.with_origin(origin),
        ),
        #[cfg(feature = "xlsb")]
        FileFormat::Xlsb => File::Xlsb(
            crate::format::xlsb::Xlsb::from_opc_package(package, options)?.with_origin(origin),
        ),
        other => return Err(not_enabled(other)),
    };
    Ok(signed_if(file, format, signed, crate::SignatureOrigin::Opc))
}

/// Build the eager handle for an ODF package detection prepared.
#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn odf_from_prepared(
    prepared: litchi_odf_common::PreparedPackage,
    options: &OpenOptions,
) -> Result<File> {
    Ok(match prepared.format() {
        #[cfg(feature = "odt")]
        FileFormat::Odt => File::Odt(crate::format::odt::Odt::from_prepared(prepared, options)?),
        #[cfg(feature = "oth")]
        FileFormat::Oth => File::Oth(crate::format::odt::Odt::from_prepared(prepared, options)?),
        #[cfg(feature = "odm")]
        FileFormat::Odm => File::Odm(crate::format::odt::Odt::from_prepared(prepared, options)?),
        #[cfg(odf_objects)]
        FileFormat::Odg | FileFormat::Odc | FileFormat::Odi | FileFormat::Odb | FileFormat::Odf
            if crate::read::Adapter::of(prepared.format()).is_some() =>
        {
            crate::format::odf_tree::OdfTree::from_prepared(prepared, options)?.into_file()
        },
        #[cfg(feature = "ods")]
        FileFormat::Ods => File::Ods(crate::format::ods::Ods::from_prepared(prepared, options)?),
        #[cfg(feature = "odp")]
        FileFormat::Odp => File::Odp(crate::format::odp::Odp::from_prepared(prepared, options)?),
        other => return Err(not_enabled(other)),
    })
}

/// Refuse a flat (single XML document) OpenDocument input: the read
/// handles open packaged ODF only, as the facade's did.
#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
fn flat_odf_unsupported(format: FileFormat) -> Error {
    Error::new(
        crate::Code::UnsupportedCapability,
        format!("flat OpenDocument {format:?} is detected, but only packaged OpenDocument opens"),
    )
    .with_format(format)
}

/// Unlock a container while carrying the origin captured by a path opener.
/// Byte-opened locked handles pass `None`; path-opened DOCX wrappers pass the
/// stable path they retained in their locked state.
#[cfg(feature = "encryption")]
pub(crate) fn unlock_with_origin(
    container: crate::locked::Container,
    bytes: Vec<u8>,
    password: &str,
    format: Option<FileFormat>,
    options: &OpenOptions,
    origin: Option<Arc<Path>>,
) -> Result<File> {
    use crate::locked::Container;

    match container {
        Container::Ooxml => unlock_ooxml_with_origin(bytes, password, format, options, origin),
        Container::Doc | Container::Ppt | Container::Xls => {
            let file = cfb::unlock(container, bytes, password, options)?;
            #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
            if let Some(origin) = origin {
                return Ok(attach_origin(file, origin));
            }
            Ok(file)
        },
    }
}

/// Decrypt an OOXML package with an optional path origin for DOCX.
#[cfg(feature = "encryption")]
fn unlock_ooxml_with_origin(
    bytes: Vec<u8>,
    password: &str,
    format: Option<FileFormat>,
    options: &OpenOptions,
    origin: Option<Arc<Path>>,
) -> Result<File> {
    #[cfg(not(any(feature = "docx", feature = "xlsx", feature = "pptx")))]
    let _ = &origin;
    let opened = litchi_crypto::ooxml::open(bytes, password, &options.ooxml_encryption_limits)?;
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    {
        let limits = options.ooxml_limits.unwrap_or_default();
        let detected = crate::detect::ooxml::try_detect_zip_format(opened.bytes(), &limits)?;
        let Some(detected) = detected else {
            return Err(Error::unsupported_format(None));
        };
        let _ = format;
        match detected {
            #[cfg(feature = "docx")]
            FileFormat::Docx => {
                let package = litchi_docx::Package::from_decrypted(
                    opened,
                    litchi_docx::OpenOptions::default()
                        .source_limits(limits)
                        .encryption_limits(options.ooxml_encryption_limits),
                )?;
                let signed = package.is_signed();
                let docx = crate::format::docx::Docx::from_owned_package_with_origin(
                    package, options, origin,
                )?;
                Ok(signed_if(
                    File::Docx(docx),
                    FileFormat::Docx,
                    signed,
                    crate::SignatureOrigin::Opc,
                ))
            },
            #[cfg(feature = "pptx")]
            FileFormat::Pptx => {
                let package = litchi_pptx::Package::from_decrypted(
                    opened,
                    litchi_pptx::OpenOptions::default()
                        .source_limits(limits)
                        .encryption_limits(options.ooxml_encryption_limits),
                )?;
                let signed = package.opc()?.is_signed();
                let pptx = crate::format::pptx::Pptx::from_package(package, options)?.with_origin(origin);
                Ok(signed_if(
                    File::Pptx(pptx),
                    FileFormat::Pptx,
                    signed,
                    crate::SignatureOrigin::Opc,
                ))
            },
            #[cfg(feature = "xlsx")]
            FileFormat::Xlsx => {
                let package = litchi_xlsx::Package::from_decrypted(
                    opened,
                    litchi_xlsx::OpenOptions::new()
                        .source_limits(limits)
                        .encryption_limits(options.ooxml_encryption_limits),
                )?;
                let signed = package.is_signed();
                let xlsx = crate::format::xlsx::Xlsx::from_package(package, options)?.with_origin(origin);
                Ok(signed_if(
                    File::Xlsx(xlsx),
                    FileFormat::Xlsx,
                    signed,
                    crate::SignatureOrigin::Opc,
                ))
            },
            FileFormat::Xlsb => Err(Error::new(
                crate::Code::UnsupportedCapability,
                "an encrypted XLSB package cannot be decrypted: the XLSB reader has no decryption support",
            )
            .with_format(FileFormat::Xlsb)),
            other => Err(not_enabled(other)),
        }
    }
    #[cfg(not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")))]
    {
        drop(opened);
        let _ = options;
        Err(format.map_or_else(
            || {
                Error::unsupported_format(None)
                    .with_suggestion("enable the `docx`, `pptx` or `xlsx` cargo feature")
            },
            not_enabled,
        ))
    }
}
