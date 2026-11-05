//! The compound-file arm of the openers: classify a CFB by its host stream,
//! recognize encryption and signatures from clear streams, and adopt the
//! legacy owner each format keeps.
//!
//! A compound file with `EncryptionInfo` and `EncryptedPackage` is an
//! encrypted OOXML package ([MS-OFFCRYPTO] 2.3.4), even when it also holds a
//! legacy host stream: Word 2007 writes a stub `WordDocument` stream next
//! to the encrypted package (record 0341). Otherwise the host-stream
//! precedence is the facade's: `WordDocument`, then `PowerPoint Document`
//! or `Current User`, then `Workbook` or `Book`. Binary DOC, PPT and XLS
//! encryption never uses an `EncryptedPackage` stream.

use std::io::Cursor;
#[cfg(any(unix, windows))]
use std::path::Path;
#[cfg(any(unix, windows))]
use std::sync::Arc;

use litchi_cfb::{OleError, OleFile, SharedOleFile};
use litchi_core::FileFormat;

#[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
use crate::format::LegacyOpen;
use crate::locked::{Container, LockedSource};
use crate::{EncryptionKind, Error, File, Locked, OpenOptions, Result};

/// What a compound file hosts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Host {
    Doc,
    Ppt,
    Xls,
    EncryptedOoxml,
    Unknown,
}

/// Read access to a parsed compound file, in memory or positional.
trait Cfb {
    fn has(&self, path: &[&str]) -> bool;
    /// Fill `output` from `path` at `offset`; `false` when the range cannot
    /// be read.
    fn read(&mut self, path: &[&str], offset: u64, output: &mut [u8]) -> bool;
}

impl Cfb for OleFile<Cursor<Vec<u8>>> {
    fn has(&self, path: &[&str]) -> bool {
        self.exists(path)
    }

    fn read(&mut self, path: &[&str], offset: u64, output: &mut [u8]) -> bool {
        self.read_stream_range(path, offset, output).is_ok()
    }
}

impl Cfb for &SharedOleFile {
    fn has(&self, path: &[&str]) -> bool {
        self.exists(path)
    }

    fn read(&mut self, path: &[&str], offset: u64, output: &mut [u8]) -> bool {
        self.read_stream_range(path, offset, output).is_ok()
    }
}

fn host(cfb: &impl Cfb) -> Host {
    if cfb.has(&["EncryptionInfo"]) && cfb.has(&["EncryptedPackage"]) {
        Host::EncryptedOoxml
    } else if cfb.has(&["WordDocument"]) {
        Host::Doc
    } else if cfb.has(&["PowerPoint Document"]) || cfb.has(&["Current User"]) {
        Host::Ppt
    } else if cfb.has(&["Workbook"]) || cfb.has(&["Book"]) {
        Host::Xls
    } else {
        Host::Unknown
    }
}

/// Whether the container carries a legacy or XML digital signature
/// (`_signatures`, `_xmlsignatures`). Nothing is verified.
fn signed(cfb: &impl Cfb) -> bool {
    cfb.has(&["_signatures"]) || cfb.has(&["_xmlsignatures"])
}

/// Container errors that mean "not a usable compound file" rather than an
/// I/O, limit or source-change failure: the input is not an Office file, as
/// the facade's detectors reported it.
fn soft(error: &OleError) -> bool {
    matches!(
        error,
        OleError::InvalidFormat(_)
            | OleError::InvalidData(_)
            | OleError::NotOleFile
            | OleError::CorruptedFile(_)
            | OleError::StreamNotFound
    )
}

/// Classify an `EncryptionVersionInfo` header ([MS-OFFCRYPTO] 2.3.4.1, and
/// the legacy RC4 headers that share its layout).
fn version_kind(header: [u8; 4], legacy: bool) -> EncryptionKind {
    let major = u16::from_le_bytes([header[0], header[1]]);
    let minor = u16::from_le_bytes([header[2], header[3]]);
    match (major, minor) {
        (4, 4) if !legacy => EncryptionKind::Agile,
        (2..=4, 2) if !legacy => EncryptionKind::Standard,
        (1, 1) if legacy => EncryptionKind::BinaryRc4,
        (2..=4, 2) => EncryptionKind::CryptoApi,
        _ => EncryptionKind::Unknown,
    }
}

/// The scheme of an encrypted OOXML package, from its `EncryptionInfo`
/// version.
fn ooxml_kind(cfb: &mut impl Cfb) -> EncryptionKind {
    let mut header = [0_u8; 4];
    if cfb.read(&["EncryptionInfo"], 0, &mut header) {
        version_kind(header, false)
    } else {
        EncryptionKind::Unknown
    }
}

/// Translate the facade's retained input ceiling to the low-level eager CFB
/// profile. The shared source route already clamps this same value before
/// parsing; eager and unlock paths must not silently revert to the 2 GiB
/// parser default when a caller selected a larger v4 input ceiling.
fn ole_limits(options: &OpenOptions) -> Result<litchi_cfb::OleFileLimits> {
    let ceiling = super::input_ceiling(options).min(litchi_cfb::OleFileLimits::MAX_INPUT_BYTES);
    Ok(litchi_cfb::OleFileLimits::new(ceiling)?)
}

/// The scheme of an encrypted Word document: the FIB's `fObfuscated` flag,
/// otherwise the encryption header at the start of its table stream.
#[cfg(feature = "doc")]
fn doc_kind(cfb: &mut impl Cfb) -> EncryptionKind {
    const F_WHICH_TBL_STM: u16 = 0x0200;
    const F_OBFUSCATED: u16 = 0x8000;

    let mut fib = [0_u8; 12];
    if !cfb.read(&["WordDocument"], 0, &mut fib) {
        return EncryptionKind::Unknown;
    }
    let flags = u16::from_le_bytes([fib[10], fib[11]]);
    if flags & F_OBFUSCATED != 0 {
        return EncryptionKind::XorObfuscation;
    }
    let table = if flags & F_WHICH_TBL_STM != 0 {
        "1Table"
    } else {
        "0Table"
    };
    let mut header = [0_u8; 4];
    if cfb.read(&[table], 0, &mut header) {
        version_kind(header, true)
    } else {
        EncryptionKind::Unknown
    }
}

/// The scheme of an encrypted Excel workbook, from its `FILEPASS` record,
/// which follows the globals' `BOF` and optional `WRITEPROT`.
#[cfg(feature = "xls")]
fn xls_kind(cfb: &mut impl Cfb) -> EncryptionKind {
    const FILEPASS: u16 = 0x002F;
    const EOF: u16 = 0x000A;
    const MAX_RECORDS: usize = 16;

    let stream: &[&str] = if cfb.has(&["Workbook"]) {
        &["Workbook"]
    } else {
        &["Book"]
    };
    let mut offset = 0_u64;
    for _ in 0..MAX_RECORDS {
        let mut header = [0_u8; 4];
        if !cfb.read(stream, offset, &mut header) {
            break;
        }
        let sid = u16::from_le_bytes([header[0], header[1]]);
        let len = u16::from_le_bytes([header[2], header[3]]);
        if sid == FILEPASS {
            let mut kind = [0_u8; 2];
            if !cfb.read(stream, offset + 4, &mut kind) {
                break;
            }
            return match u16::from_le_bytes(kind) {
                0 => EncryptionKind::XorObfuscation,
                1 => {
                    let mut version = [0_u8; 4];
                    if cfb.read(stream, offset + 6, &mut version) {
                        version_kind(version, true)
                    } else {
                        EncryptionKind::Unknown
                    }
                },
                _ => EncryptionKind::Unknown,
            };
        }
        if sid == EOF {
            break;
        }
        offset += 4 + u64::from(len);
    }
    EncryptionKind::Unknown
}

/// Refuse a compound file an OOXML file name (`ooxml_hint`) calls a package
/// unless it hosts an encrypted OOXML package: the facade's suffix preflight
/// refused every compound file under such a name. `host` is `None` for a
/// container that could not be parsed.
fn check_ooxml_hint(ooxml_hint: Option<FileFormat>, host: Option<Host>) -> Result<()> {
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    if ooxml_hint.is_some() && host != Some(Host::EncryptedOoxml) {
        return Err(super::ooxml_suffix_mismatch());
    }
    let _ = (ooxml_hint, host);
    Ok(())
}

/// The handle of an encrypted input, keeping its in-memory source.
fn locked(
    container: Container,
    format: Option<FileFormat>,
    encryption: EncryptionKind,
    ole: OleFile<Cursor<Vec<u8>>>,
    options: &OpenOptions,
) -> File {
    File::Locked(Locked::new(
        container,
        format,
        encryption,
        LockedSource::Owned(ole.into_inner().into_inner()),
        options,
    ))
}

/// Open an in-memory compound file. `ooxml_hint` is the format a file name
/// named, reported for an encrypted OOXML package.
pub(super) fn open_bytes(
    bytes: Vec<u8>,
    options: &OpenOptions,
    ooxml_hint: Option<FileFormat>,
) -> Result<File> {
    let limits = ole_limits(options)?;
    let mut ole = match OleFile::open(Cursor::new(bytes), limits) {
        Ok(ole) => ole,
        Err(error) if soft(&error) => {
            check_ooxml_hint(ooxml_hint, None)?;
            return Err(Error::unsupported_format(None));
        },
        Err(error) => return Err(error.into()),
    };
    let host = host(&ole);
    check_ooxml_hint(ooxml_hint, Some(host))?;
    match host {
        Host::EncryptedOoxml => {
            let kind = ooxml_kind(&mut ole);
            Ok(locked(Container::Ooxml, ooxml_hint, kind, ole, options))
        },
        Host::Unknown => Err(Error::unsupported_format(None)),
        legacy => open_legacy(legacy, ole, options, None),
    }
}

/// Open the legacy document an in-memory compound file hosts, with
/// `password` when an unlock supplies one.
#[cfg_attr(
    not(all(feature = "doc", feature = "ppt", feature = "xls")),
    allow(
        unused_variables,
        clippy::needless_pass_by_value,
        reason = "a build without a legacy format refuses its host instead of opening it"
    )
)]
fn open_legacy(
    host: Host,
    ole: OleFile<Cursor<Vec<u8>>>,
    options: &OpenOptions,
    password: Option<&str>,
) -> Result<File> {
    let is_signed = signed(&ole);
    match host {
        #[cfg(feature = "doc")]
        Host::Doc => Ok(
            match crate::format::doc::Doc::from_ole_file(ole, options, password)? {
                LegacyOpen::Opened(doc) => super::signed_if(
                    File::Doc(doc),
                    FileFormat::Doc,
                    is_signed,
                    crate::SignatureOrigin::Cfb,
                ),
                LegacyOpen::PasswordRequired(mut ole) => {
                    let kind = doc_kind(ole.as_mut());
                    locked(Container::Doc, Some(FileFormat::Doc), kind, *ole, options)
                },
            },
        ),
        #[cfg(feature = "ppt")]
        Host::Ppt => Ok(
            match crate::format::ppt::Ppt::from_ole_file(ole, options, password)? {
                LegacyOpen::Opened(ppt) => super::signed_if(
                    File::Ppt(ppt),
                    FileFormat::Ppt,
                    is_signed,
                    crate::SignatureOrigin::Cfb,
                ),
                // PowerPoint's only password-to-open scheme is RC4
                // `CryptoAPI` ([MS-PPT] 2.3.7).
                LegacyOpen::PasswordRequired(ole) => locked(
                    Container::Ppt,
                    Some(FileFormat::Ppt),
                    EncryptionKind::CryptoApi,
                    *ole,
                    options,
                ),
            },
        ),
        #[cfg(feature = "xls")]
        Host::Xls => Ok(
            match crate::format::xls::Xls::from_ole_file(ole, options, password)? {
                LegacyOpen::Opened(xls) => super::signed_if(
                    File::Xls(xls),
                    FileFormat::Xls,
                    is_signed,
                    crate::SignatureOrigin::Cfb,
                ),
                LegacyOpen::PasswordRequired(mut ole) => {
                    let kind = xls_kind(ole.as_mut());
                    locked(Container::Xls, Some(FileFormat::Xls), kind, *ole, options)
                },
            },
        ),
        #[cfg(not(feature = "doc"))]
        Host::Doc => Err(super::not_enabled(FileFormat::Doc)),
        #[cfg(not(feature = "ppt"))]
        Host::Ppt => Err(super::not_enabled(FileFormat::Ppt)),
        #[cfg(not(feature = "xls"))]
        Host::Xls => Err(super::not_enabled(FileFormat::Xls)),
        Host::EncryptedOoxml | Host::Unknown => Err(Error::unsupported_format(None)),
    }
}

/// Open a positional compound file.
///
/// A file whose name says it is a Word document is read once and parsed in
/// memory, the facade's document route; the extension only picks the route,
/// and the content still decides the format. Every other compound file is
/// parsed through the positional source, which the PowerPoint and Excel
/// owners keep reading on demand.
#[cfg(any(unix, windows))]
pub(super) fn open_path(
    path: &Path,
    source: Arc<dyn litchi_core::ReadAt>,
    version: litchi_core::SourceVersion,
    options: &OpenOptions,
) -> Result<File> {
    use crate::detect::detected::ensure_path_source_current;

    let ooxml_hint = super::ooxml_format_for_extension(path);
    let ceiling = super::input_ceiling(options);
    if names_word_document(path) {
        let bytes = super::read_whole_source(source.as_ref(), version, ceiling)?;
        return open_bytes(bytes, options, ooxml_hint);
    }
    let shared_limits = litchi_cfb::SharedOleFileLimits::new(
        options
            .max_input_bytes
            .unwrap_or(litchi_cfb::SharedOleFileLimits::DEFAULT_MAX_INPUT_BYTES)
            .min(litchi_cfb::SharedOleFileLimits::MAX_INPUT_BYTES),
    )?;
    let shared = match SharedOleFile::open(Arc::clone(&source), shared_limits) {
        Ok(shared) => shared,
        Err(error) if soft(&error) => {
            ensure_path_source_current(source.as_ref(), version)?;
            check_ooxml_hint(ooxml_hint, None)?;
            return Err(Error::unsupported_format(None));
        },
        Err(error) => return Err(error.into()),
    };
    let host = host(&&shared);
    check_ooxml_hint(ooxml_hint, Some(host))?;
    let file = match host {
        // A Word document by another name: the in-memory document route.
        Host::Doc => {
            drop(shared);
            let bytes = super::read_whole_source(source.as_ref(), version, ceiling)?;
            return open_bytes(bytes, options, ooxml_hint);
        },
        Host::Ppt => ppt_source(shared, &source, version, options)?,
        Host::Xls => xls_source(shared, &source, version, options)?,
        Host::EncryptedOoxml => {
            let kind = ooxml_kind(&mut &shared);
            File::Locked(Locked::new_with_origin(
                Container::Ooxml,
                ooxml_hint,
                kind,
                LockedSource::Positional {
                    source: Arc::clone(&source),
                    version,
                },
                options,
                Some(crate::open::path::stable_origin(path)?),
            ))
        },
        Host::Unknown => return Err(Error::unsupported_format(None)),
    };
    ensure_path_source_current(source.as_ref(), version)?;
    Ok(file)
}

/// Whether a file name names a legacy Word document or template.
#[cfg(any(unix, windows))]
fn names_word_document(path: &Path) -> bool {
    path.extension()
        .and_then(std::ffi::OsStr::to_str)
        .is_some_and(|extension| {
            ["doc", "dot", "wbk"]
                .iter()
                .any(|known| extension.eq_ignore_ascii_case(known))
        })
}

/// The facade's PowerPoint path route: the positional package over the
/// compound file just parsed.
#[cfg(any(unix, windows))]
#[cfg_attr(
    not(feature = "ppt"),
    allow(
        clippy::needless_pass_by_value,
        unused_variables,
        reason = "without `ppt` the host is refused"
    )
)]
fn ppt_source(
    shared: SharedOleFile,
    source: &Arc<dyn litchi_core::ReadAt>,
    version: litchi_core::SourceVersion,
    options: &OpenOptions,
) -> Result<File> {
    #[cfg(feature = "ppt")]
    {
        let is_signed = signed(&&shared);
        let Some(package) =
            litchi_ppt::SourceBackedPackage::from_shared_if_powerpoint_with_options(
                shared,
                litchi_ppt::OpenOptions::new()
                    .source_limits(options.ppt_source_limits.unwrap_or_default()),
            )?
        else {
            // A `Current User` stream without a PowerPoint document: the
            // byte route reports what the reader makes of it.
            let bytes =
                super::read_whole_source(source.as_ref(), version, super::input_ceiling(options))?;
            return open_bytes(bytes, options, None);
        };
        match crate::format::ppt::Ppt::from_source_package(package, options) {
            Ok(ppt) => Ok(super::signed_if(
                File::Ppt(ppt),
                FileFormat::Ppt,
                is_signed,
                crate::SignatureOrigin::Cfb,
            )),
            Err(error)
                if matches!(
                    error.leaf(),
                    crate::error::Source::Ppt(leaf) if crate::format::ppt::password_required(leaf)
                ) =>
            {
                Ok(File::Locked(Locked::new(
                    Container::Ppt,
                    Some(FileFormat::Ppt),
                    EncryptionKind::CryptoApi,
                    LockedSource::Positional {
                        source: Arc::clone(source),
                        version,
                    },
                    options,
                )))
            },
            Err(error) => Err(error),
        }
    }
    #[cfg(not(feature = "ppt"))]
    Err(super::not_enabled(FileFormat::Ppt))
}

/// The facade's Excel path route: the positional workbook over the compound
/// file just parsed, falling back to the eager in-memory reader for the
/// workbooks the positional owner does not read (BIFF5 and older).
#[cfg(any(unix, windows))]
#[cfg_attr(
    not(feature = "xls"),
    allow(
        clippy::needless_pass_by_value,
        unused_variables,
        reason = "without `xls` the host is refused"
    )
)]
fn xls_source(
    shared: SharedOleFile,
    source: &Arc<dyn litchi_core::ReadAt>,
    version: litchi_core::SourceVersion,
    options: &OpenOptions,
) -> Result<File> {
    #[cfg(feature = "xls")]
    {
        use litchi_xls::SourceBackedError;

        let is_signed = signed(&&shared);
        let shared = Arc::new(shared);
        match litchi_xls::raw::source_backed_workbook_from_shared_ole_file(
            Arc::clone(&shared),
            litchi_xls::SourceBackedLimits::default().with_max_input_bytes(
                options
                    .max_input_bytes
                    .unwrap_or(super::NEUTRAL_MAX_INPUT_BYTES)
                    .min(litchi_cfb::SharedOleFileLimits::MAX_INPUT_BYTES),
            ),
        ) {
            Ok(workbook) => {
                let observed = workbook.source_version()?;
                if observed != version {
                    return Err(litchi_core::Error::SourceChanged {
                        expected: version,
                        observed,
                    }
                    .into());
                }
                Ok(super::signed_if(
                    File::Xls(crate::format::xls::Xls::from_source(workbook)?),
                    FileFormat::Xls,
                    is_signed,
                    crate::SignatureOrigin::Cfb,
                ))
            },
            Err(SourceBackedError::EncryptedUnsupported) => {
                let kind = xls_kind(&mut &*shared);
                Ok(File::Locked(Locked::new(
                    Container::Xls,
                    Some(FileFormat::Xls),
                    kind,
                    LockedSource::Positional {
                        source: Arc::clone(source),
                        version,
                    },
                    options,
                )))
            },
            Err(
                SourceBackedError::WorkbookStreamMissing
                | SourceBackedError::UnsupportedBiffVersion(_),
            ) => {
                drop(shared);
                let bytes = super::read_whole_source(
                    source.as_ref(),
                    version,
                    super::input_ceiling(options),
                )?;
                open_bytes(bytes, options, None)
            },
            Err(error) => Err(error.into()),
        }
    }
    #[cfg(not(feature = "xls"))]
    Err(super::not_enabled(FileFormat::Xls))
}

/// Open the content of a locked legacy document with `password`.
#[cfg(feature = "encryption")]
pub(super) fn unlock(
    container: Container,
    bytes: Vec<u8>,
    password: &str,
    options: &OpenOptions,
) -> Result<File> {
    let ole = OleFile::open(Cursor::new(bytes), ole_limits(options)?)?;
    let host = match container {
        Container::Doc => Host::Doc,
        Container::Ppt => Host::Ppt,
        Container::Xls => Host::Xls,
        Container::Ooxml => Host::EncryptedOoxml,
    };
    match open_legacy(host, ole, options, Some(password))? {
        // The leaf still asks for a password it was given: refuse as locked
        // rather than hand back a second `Locked`.
        File::Locked(locked) => Err(locked.read_error()),
        file => Ok(file),
    }
}

#[cfg(test)]
mod tests {
    use super::ole_limits;
    use crate::OpenOptions;

    #[test]
    fn eager_cfb_profile_follows_explicit_v4_input_ceiling() {
        let options = OpenOptions::new().max_input_bytes(3 * 1024 * 1024 * 1024);
        let limits = ole_limits(&options).expect("valid CFB input profile");
        assert_eq!(limits.max_input_bytes(), 3 * 1024 * 1024 * 1024);
    }
}
