//! Filesystem publication policy for the unified [`File`] surface.
//!
//! The policy types in this module only choose a destination and a leaf
//! publication policy.  Format writers remain in their owning format module;
//! the unified layer does not implement a second archive writer or its own
//! rename and synchronization sequence.

use std::fmt;
use std::path::{Path, PathBuf};

#[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
use litchi_cfb::atomic::{self, Overwrite as CfbOverwrite};

use litchi_core::{Durability, FileFormat};

use crate::File;
use crate::error::{Code, Error, Result};

/// Where [`File::save_with`] publishes a snapshot.
#[derive(Clone, Copy, Default)]
#[non_exhaustive]
pub enum SaveTo<'a> {
    /// The path retained by a path-backed open.
    #[default]
    Origin,
    /// An explicit destination.  The file's retained origin is unchanged.
    Path(&'a Path),
}

impl<'a> From<&'a Path> for SaveTo<'a> {
    fn from(path: &'a Path) -> Self {
        Self::Path(path)
    }
}

impl<'a> From<&'a PathBuf> for SaveTo<'a> {
    fn from(path: &'a PathBuf) -> Self {
        Self::Path(path)
    }
}

impl<'a> From<&'a str> for SaveTo<'a> {
    fn from(path: &'a str) -> Self {
        Self::Path(Path::new(path))
    }
}

impl fmt::Debug for SaveTo<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Origin => formatter.write_str("SaveTo::Origin"),
            // Destinations are caller-authored input; keep Debug structural
            // and content-free.
            Self::Path(_) => formatter.write_str("SaveTo::Path(..)"),
        }
    }
}

/// What a save does when its destination already exists.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Overwrite {
    /// Replace the destination atomically.  This is the default.
    #[default]
    Replace,
    /// Refuse an existing destination without changing it.
    Refuse,
}

/// Options for [`File::save_with`].
///
/// The defaults are full durability and replacement of an existing file.
/// A password, when present, is owned by this options value and borrowed only
/// while a sensitive package is re-encrypted; it is never copied into a clone
/// or formatted diagnostic.
#[derive(Debug, Default)]
#[non_exhaustive]
pub struct SaveOptions {
    /// Synchronization policy for the leaf's atomic publication.
    pub(crate) durability: Durability,
    /// Existing-destination policy.
    pub(crate) overwrite: Overwrite,
    /// Password for re-encrypting a sensitive OOXML package.
    #[cfg(feature = "encryption")]
    pub(crate) reencrypt: Option<crate::Password>,
}

impl SaveOptions {
    /// Construct the default save policy.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            durability: Durability::Full,
            overwrite: Overwrite::Replace,
            #[cfg(feature = "encryption")]
            reencrypt: None,
        }
    }

    /// Set the leaf synchronization policy.
    #[must_use]
    pub const fn durability(mut self, durability: Durability) -> Self {
        self.durability = durability;
        self
    }

    /// Set the existing-destination policy.
    #[must_use]
    pub const fn overwrite(mut self, overwrite: Overwrite) -> Self {
        self.overwrite = overwrite;
        self
    }

    /// Supply the password used to re-encrypt a sensitive package with its
    /// retained encryption profile.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub fn reencrypt(mut self, password: crate::Password) -> Self {
        self.reencrypt = Some(password);
        self
    }
}

/// The result of a successful path save.
#[derive(Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct Saved {
    /// Number of bytes accepted by the leaf writer.
    pub bytes: u64,
    /// Destination supplied to the save operation.
    pub path: PathBuf,
}

impl fmt::Debug for Saved {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Saved")
            .field("bytes", &self.bytes)
            .finish_non_exhaustive()
    }
}

impl File {
    /// Save this snapshot to its retained origin path.
    ///
    /// The destination is selected from the path retained by the opener. The
    /// leaf publishes through its atomic writer using full durability and
    /// replacement as the defaults.
    ///
    /// # Errors
    ///
    /// Returns [`Code::FileRequired`] when this handle has no retained path,
    /// or a typed source, destination, policy, leaf or I/O error.
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # fn main() -> litchi_ops::Result<()> {
    /// let file = litchi_ops::open("report.docx")?;
    /// let _saved = file.save()?;
    /// # Ok(())
    /// # }
    /// # #[cfg(not(feature = "docx"))]
    /// # fn main() {}
    /// ```
    pub fn save(&self) -> Result<Saved> {
        self.save_with(SaveTo::Origin, &SaveOptions::new())
    }

    /// Save this snapshot to an explicit path, leaving its origin unchanged.
    ///
    /// The destination extension must name the same format when it is a
    /// recognized office extension. Full durability and replacement are used
    /// by default.
    ///
    /// # Errors
    ///
    /// Returns a typed destination, policy, leaf or I/O error. The input
    /// handle and its retained origin remain unchanged.
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # fn main() -> litchi_ops::Result<()> {
    /// let file = litchi_ops::open("report.docx")?;
    /// let _saved = file.save_as("report-copy.docx")?;
    /// # Ok(())
    /// # }
    /// # #[cfg(not(feature = "docx"))]
    /// # fn main() {}
    /// ```
    pub fn save_as(&self, path: impl AsRef<Path>) -> Result<Saved> {
        self.save_with(SaveTo::Path(path.as_ref()), &SaveOptions::new())
    }

    /// Save this snapshot using an explicit destination and leaf policy.
    ///
    /// `Overwrite::Refuse` is enforced by the leaf's atomic no-clobber
    /// operation, including when another process creates the destination
    /// during staging. Source-backed handles are fenced immediately before
    /// publication.
    ///
    /// # Errors
    ///
    /// Returns a typed source, destination, policy, leaf or I/O error.
    /// Failures before replacement leave the destination unchanged; a leaf
    /// directory-sync failure after replacement is reported with its committed
    /// status.
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # fn main() -> litchi_ops::Result<()> {
    /// use std::path::Path;
    /// use litchi_ops::{Overwrite, SaveOptions, SaveTo};
    ///
    /// let file = litchi_ops::open("report.docx")?;
    /// let options = SaveOptions::new().overwrite(Overwrite::Refuse);
    /// let _saved = file.save_with(
    ///     SaveTo::Path(Path::new("report-copy.docx")),
    ///     &options,
    /// )?;
    /// # Ok(())
    /// # }
    /// # #[cfg(not(feature = "docx"))]
    /// # fn main() {}
    /// ```
    pub fn save_with(&self, to: SaveTo<'_>, options: &SaveOptions) -> Result<Saved> {
        // A signed wrapper delegates to the wrapped handle.  The wrapped
        // format retains its signature disposition, so an ordinary signed
        // save preserves the signature while an explicitly unsigned handle
        // reaches the leaf's removal-authorized route.
        if let Self::Signed(signed) = self {
            return signed.file().save_with(to, options);
        }

        #[cfg(not(any(
            feature = "doc",
            feature = "docx",
            feature = "ppt",
            feature = "xlsx",
            feature = "pptx",
            feature = "xls",
            feature = "xlsb",
            feature = "rtf",
            feature = "ods",
            feature = "odt",
            feature = "odp",
            odf_objects,
            feature = "iwork-adapter"
        )))]
        {
            let _ = (to, options);
            return match self {
                Self::Locked(locked) => Err(locked.read_error()),
                _ => Err(unsupported_save(self.format())),
            };
        }

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
            feature = "odt",
            feature = "odp",
            odf_objects,
            feature = "iwork-adapter"
        ))]
        {
            // Validate the handle state before looking at caller-supplied path
            // text or destination metadata. A locked or unsupported file must
            // retain precedence over extension and overwrite refusals.
            match self {
                Self::Locked(locked) => return Err(locked.read_error()),
                #[cfg(feature = "doc")]
                Self::Doc(_) => {},
                #[cfg(feature = "rtf")]
                Self::Rtf(_) => {},
                #[cfg(feature = "odt")]
                Self::Odt(_) | Self::Oth(_) | Self::Odm(_) => {},
                #[cfg(odf_objects)]
                Self::Odg(_) | Self::Odc(_) | Self::Odi(_) | Self::Odb(_) | Self::Odf(_) => {},
                #[cfg(feature = "iwork-adapter")]
                Self::Pages(_) | Self::Keynote(_) | Self::Numbers(_) => {},
                #[cfg(feature = "odp")]
                Self::Odp(_) => {},
                #[cfg(feature = "xls")]
                Self::Xls(_) => {},
                #[cfg(feature = "docx")]
                Self::Docx(_) => {},
                #[cfg(feature = "ppt")]
                Self::Ppt(_) => {},
                #[cfg(feature = "xlsx")]
                Self::Xlsx(_) => {},
                #[cfg(feature = "xlsb")]
                Self::Xlsb(_) => {},
                #[cfg(feature = "pptx")]
                Self::Pptx(_) => {},
                #[cfg(feature = "ods")]
                Self::Ods(_) => {},
                #[allow(unreachable_patterns)]
                _ => return Err(unsupported_save(self.format())),
            }

            let format = self.format();
            let destination = match to {
                SaveTo::Origin => origin_path(self, format)?,
                SaveTo::Path(path) => {
                    if let Some(format) = format {
                        same_format(path, format)?;
                    }
                    path.to_path_buf()
                },
            };

            let bytes = match self {
                #[cfg(feature = "doc")]
                Self::Doc(handle) => save_doc(handle, &destination, options)?,
                #[cfg(feature = "rtf")]
                Self::Rtf(handle) => handle.save_to(&destination, options)?,
                #[cfg(feature = "odt")]
                Self::Odt(handle) | Self::Oth(handle) | Self::Odm(handle) => {
                    handle.save_to(&destination, options)?
                },
                #[cfg(odf_objects)]
                Self::Odg(handle)
                | Self::Odc(handle)
                | Self::Odi(handle)
                | Self::Odb(handle)
                | Self::Odf(handle) => handle.save_to(&destination, options)?,
                #[cfg(feature = "iwork-adapter")]
                Self::Pages(handle) | Self::Keynote(handle) | Self::Numbers(handle) => {
                    handle.save_to(&destination, options)?
                },
                #[cfg(feature = "odp")]
                Self::Odp(handle) => handle.save_to(&destination, options)?,
                #[cfg(feature = "xls")]
                Self::Xls(handle) => save_xls(handle, &destination, options)?,
                #[cfg(feature = "docx")]
                Self::Docx(handle) => handle.save_to(&destination, options)?,
                #[cfg(feature = "ppt")]
                Self::Ppt(handle) => save_ppt(handle, &destination, options)?,
                #[cfg(feature = "xlsx")]
                Self::Xlsx(handle) => handle.save_to(&destination, options)?,
                #[cfg(feature = "xlsb")]
                Self::Xlsb(handle) => handle.save_to(&destination, options)?,
                #[cfg(feature = "pptx")]
                Self::Pptx(handle) => handle.save_to(&destination, options)?,
                #[cfg(feature = "ods")]
                Self::Ods(handle) => handle.save_to(&destination, options)?,
                #[allow(unreachable_patterns)]
                _ => return Err(unsupported_save(format)),
            };

            Ok(Saved {
                bytes,
                path: destination,
            })
        }
    }
}

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
    feature = "odt",
    feature = "odp",
    odf_objects,
    feature = "iwork-adapter"
))]
fn origin_path(file: &File, format: Option<FileFormat>) -> Result<PathBuf> {
    #[cfg(feature = "iwork-adapter")]
    if let Some(handle) = file.iwork() {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(odf_objects)]
    if let Some(handle) = file.odf_tree() {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(feature = "odt")]
    if let Some(handle) = file.odt_family() {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(feature = "odp")]
    if let File::Odp(handle) = file {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(feature = "rtf")]
    if let File::Rtf(handle) = file {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(feature = "xls")]
    if let File::Xls(handle) = file {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }

    #[cfg(feature = "doc")]
    if let File::Doc(handle) = file {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(feature = "docx")]
    if let File::Docx(handle) = file {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(feature = "ppt")]
    if let File::Ppt(handle) = file {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(feature = "xlsb")]
    if let File::Xlsb(handle) = file {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(feature = "xlsx")]
    if let File::Xlsx(handle) = file {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(feature = "pptx")]
    if let File::Pptx(handle) = file {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    #[cfg(feature = "ods")]
    if let File::Ods(handle) = file {
        return handle
            .origin()
            .map(|origin| origin.as_ref().to_path_buf())
            .ok_or_else(|| file_required(format));
    }
    Err(file_required(format))
}

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
    feature = "odt",
    feature = "odp",
    odf_objects,
    feature = "iwork-adapter"
))]
fn file_required(format: Option<FileFormat>) -> Error {
    let error = Error::new(
        Code::FileRequired,
        "the document has no path to save to; save it with `File::save_as`",
    )
    .with_suggestion("save it with `File::save_as(path)`, or write it with `File::write_to`");
    format.map_or(error.clone(), |format| error.with_format(format))
}

fn unsupported_save(format: Option<FileFormat>) -> Error {
    let error = Error::new(
        Code::UnsupportedCapability,
        "saving this format is not integrated yet",
    )
    .with_suggestion("see the format capability report and M2 follow-up");
    format.map_or(error.clone(), |format| error.with_format(format))
}

/// Publish a DOC source through the CFB sibling-temporary publisher while
/// streaming the native owner directly into the staged file. This keeps both
/// retained CFB sources and edited body snapshots on their existing
/// publication path: no facade byte buffer or second package model is
/// introduced.
#[cfg(feature = "doc")]
fn save_doc(
    handle: &crate::format::doc::Doc,
    destination: &Path,
    options: &SaveOptions,
) -> Result<u64> {
    handle.check_source()?;
    validate_doc_destination(destination)?;
    let overwrite = match options.overwrite {
        Overwrite::Replace => CfbOverwrite::Replace,
        Overwrite::Refuse => CfbOverwrite::Refuse,
    };
    atomic::save(
        destination,
        options.durability,
        overwrite,
        |staging| handle.write_to(staging),
        || handle.check_source(),
    )
}

#[cfg(feature = "doc")]
fn validate_doc_destination(destination: &Path) -> Result<()> {
    if destination.file_name().is_none() {
        return Err(Error::new(
            Code::InvalidArgument,
            "the DOC save destination must name a file",
        )
        .with_format(FileFormat::Doc));
    }
    match std::fs::symlink_metadata(destination) {
        Ok(metadata) => {
            if metadata.file_type().is_symlink() || !metadata.is_file() {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "the DOC save destination must be a regular file",
                )
                .with_format(FileFormat::Doc));
            }
            Ok(())
        },
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(Error::from(error)),
    }
}

#[cfg(feature = "xls")]
fn save_xls(
    handle: &crate::format::xls::Xls,
    destination: &Path,
    options: &SaveOptions,
) -> Result<u64> {
    handle.check_source()?;
    validate_xls_destination(destination)?;
    let overwrite = match options.overwrite {
        Overwrite::Replace => CfbOverwrite::Replace,
        Overwrite::Refuse => CfbOverwrite::Refuse,
    };
    atomic::save(
        destination,
        options.durability,
        overwrite,
        |staging| {
            let mut counted = crate::write::Counted {
                sink: staging,
                bytes: 0,
            };
            handle.write_to(&mut counted)?;
            Ok(counted.bytes)
        },
        || handle.check_source(),
    )
}

#[cfg(feature = "ppt")]
fn save_ppt(
    handle: &crate::format::ppt::Ppt,
    destination: &Path,
    options: &SaveOptions,
) -> Result<u64> {
    handle.check_source()?;
    validate_ppt_destination(destination)?;
    let overwrite = match options.overwrite {
        Overwrite::Replace => CfbOverwrite::Replace,
        Overwrite::Refuse => CfbOverwrite::Refuse,
    };
    let mut bytes = 0_u64;
    atomic::save(
        destination,
        options.durability,
        overwrite,
        |staging| {
            let mut counted = crate::write::Counted {
                sink: staging,
                bytes: 0,
            };
            handle.write_to(&mut counted)?;
            bytes = counted.bytes;
            Ok(())
        },
        || handle.check_source(),
    )?;
    Ok(bytes)
}

#[cfg(feature = "ppt")]
fn validate_ppt_destination(destination: &Path) -> Result<()> {
    if destination.file_name().is_none() {
        return Err(Error::new(
            Code::InvalidArgument,
            "the PPT save destination must name a file",
        )
        .with_format(FileFormat::Ppt));
    }
    match std::fs::symlink_metadata(destination) {
        Ok(metadata) => {
            if metadata.file_type().is_symlink() || !metadata.is_file() {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "the PPT save destination must be a regular file",
                )
                .with_format(FileFormat::Ppt));
            }
            Ok(())
        },
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(Error::from(error)),
    }
}

#[cfg(feature = "xls")]
fn validate_xls_destination(destination: &Path) -> Result<()> {
    if destination.file_name().is_none() {
        return Err(Error::new(
            Code::InvalidArgument,
            "the XLS save destination must name a file",
        )
        .with_format(FileFormat::Xls));
    }
    match std::fs::symlink_metadata(destination) {
        Ok(metadata) => {
            if metadata.file_type().is_symlink() || !metadata.is_file() {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "the XLS save destination must be a regular file",
                )
                .with_format(FileFormat::Xls));
            }
            Ok(())
        },
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(Error::from(error)),
    }
}

/// Refuse a destination whose extension names a different format.  Flavour
/// extensions of the same format, unknown extensions and no extension are
/// accepted because content, rather than a filename, chooses the format.
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
    feature = "odt",
    feature = "odp",
    odf_objects,
    feature = "iwork-adapter"
))]
fn same_format(path: &Path, format: FileFormat) -> Result<()> {
    let Some(extension) = path.extension().and_then(|extension| extension.to_str()) else {
        return Ok(());
    };
    if let Some(named) = format_of_extension(extension) {
        if named != format {
            return Err(Error::new(
                Code::InvalidArgument,
                "the destination extension names another format; a save keeps the document's own format",
            )
            .with_format(format)
            .with_suggestion("conversion is explicit: use `export` (not in M2)"));
        }
    }
    Ok(())
}

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
    feature = "odt",
    feature = "odp",
    odf_objects,
    feature = "iwork-adapter"
))]
fn format_of_extension(extension: &str) -> Option<FileFormat> {
    const EXTENSIONS: &[(&str, FileFormat)] = &[
        ("doc", FileFormat::Doc),
        ("docm", FileFormat::Docx),
        ("docx", FileFormat::Docx),
        ("dot", FileFormat::Doc),
        ("dotm", FileFormat::Docx),
        ("dotx", FileFormat::Docx),
        ("key", FileFormat::Keynote),
        ("numbers", FileFormat::Numbers),
        ("odb", FileFormat::Odb),
        ("odc", FileFormat::Odc),
        ("odf", FileFormat::Odf),
        ("odg", FileFormat::Odg),
        ("odi", FileFormat::Odi),
        ("odm", FileFormat::Odm),
        ("odp", FileFormat::Odp),
        ("ods", FileFormat::Ods),
        ("odt", FileFormat::Odt),
        ("otc", FileFormat::Odc),
        ("otg", FileFormat::Odg),
        ("oth", FileFormat::Oth),
        ("oti", FileFormat::Odi),
        ("otm", FileFormat::Odm),
        ("otp", FileFormat::Odp),
        ("ots", FileFormat::Ods),
        ("ott", FileFormat::Odt),
        ("pages", FileFormat::Pages),
        ("pot", FileFormat::Ppt),
        ("potm", FileFormat::Pptx),
        ("potx", FileFormat::Pptx),
        ("pps", FileFormat::Ppt),
        ("ppsm", FileFormat::Pptx),
        ("ppsx", FileFormat::Pptx),
        ("ppt", FileFormat::Ppt),
        ("pptm", FileFormat::Pptx),
        ("pptx", FileFormat::Pptx),
        ("rtf", FileFormat::Rtf),
        ("xls", FileFormat::Xls),
        ("xlsb", FileFormat::Xlsb),
        ("xlsm", FileFormat::Xlsx),
        ("xlsx", FileFormat::Xlsx),
        ("xlt", FileFormat::Xls),
        ("xltm", FileFormat::Xlsx),
        ("xltx", FileFormat::Xlsx),
    ];
    EXTENSIONS
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(extension))
        .map(|&(_, format)| format)
}
