//! The [`Locked`] state of [`File`](crate::File): an encrypted input that
//! is neither readable nor editable until it is unlocked (ADR 0006
//! `Locked<T>`; record 0003 §2.6).

use std::fmt;
use std::sync::Arc;

use litchi_core::FileFormat;

use crate::error::Error;
use crate::{Family, OpenOptions};

/// The encryption scheme of a [`Locked`] input, as classified without a
/// password.
///
/// OOXML kinds come from the `EncryptionInfo` version; legacy kinds come
/// from the format's clear encryption header (the `FILEPASS` record, the FIB
/// and table-stream header, or the PowerPoint `CryptoAPI` container).
/// [`Self::Unknown`] means the input was recognized as encrypted but its
/// scheme was not classified.
///
/// ```
/// use litchi_ops::EncryptionKind;
///
/// assert_ne!(EncryptionKind::Agile, EncryptionKind::Standard);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum EncryptionKind {
    /// OOXML Agile Encryption (`EncryptionInfo` version 4.4).
    Agile,
    /// OOXML Standard Encryption (`EncryptionInfo` versions 2.2 to 4.2).
    Standard,
    /// Legacy Office binary RC4 encryption (not `CryptoAPI`).
    BinaryRc4,
    /// Legacy RC4 `CryptoAPI` encryption.
    CryptoApi,
    /// Legacy XOR obfuscation.
    XorObfuscation,
    /// Encrypted, scheme not classified.
    Unknown,
}

/// An encrypted input: recognized, but neither readable nor editable until
/// it is unlocked with its password (ADR 0006 `Locked<T>`).
///
/// Reads and family views of a [`File::Locked`](crate::File::Locked) return
/// [`Code::DocumentProtected`](crate::Code::DocumentProtected). The handle
/// keeps its source and the options it was opened with, so the consuming
/// `Locked::unlock` (with the `encryption` cargo feature) opens the content
/// without opening the input again. Credentials never pass through
/// [`OpenOptions`]: `unlock` takes the password itself.
///
/// `Locked` is a cheap-to-clone handle. Its [`Debug`] output names the inner
/// format and encryption kind only; it never prints source bytes or key
/// material.
///
/// ```no_run
/// use litchi_ops::File;
///
/// # #[cfg(feature = "docx")]
/// # fn main() -> litchi_ops::Result<()> {
/// if let File::Locked(locked) = litchi_ops::open("secret.docx")? {
///     println!("{:?} encrypted with {:?}", locked.format(), locked.encryption());
/// }
/// # Ok(())
/// # }
/// # #[cfg(not(feature = "docx"))]
/// # fn main() {}
/// ```
#[derive(Clone)]
pub struct Locked(Arc<LockedState>);

/// Shared state behind a [`Locked`] handle.
#[allow(
    dead_code,
    reason = "the source, container and options are read only by `unlock`, which needs the `encryption` feature"
)]
pub(crate) struct LockedState {
    container: Container,
    format: Option<FileFormat>,
    encryption: EncryptionKind,
    source: LockedSource,
    options: OpenOptions,
    /// Stable path origin, when the encrypted container was opened by path.
    /// It is metadata only and is intentionally omitted from `Debug`.
    origin: Option<Arc<std::path::Path>>,
}

/// The kind of container an encrypted input arrived in.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    dead_code,
    reason = "a build constructs only the containers its formats read, and `unlock` matches them"
)]
pub(crate) enum Container {
    /// An OOXML package encrypted into a compound file.
    Ooxml,
    /// A legacy Word document with password-to-open encryption.
    Doc,
    /// A legacy PowerPoint presentation with password-to-open encryption.
    Ppt,
    /// A legacy Excel workbook with password-to-open encryption.
    Xls,
}

/// The source a [`Locked`] handle keeps for its unlock.
#[allow(
    dead_code,
    reason = "read only by `unlock`, which needs the `encryption` feature"
)]
pub(crate) enum LockedSource {
    /// The in-memory input, moved into the handle without a copy.
    Owned(Vec<u8>),
    /// The positional file source, pinned to the revision opened.
    #[cfg(any(unix, windows))]
    Positional {
        source: Arc<dyn litchi_core::ReadAt>,
        version: litchi_core::SourceVersion,
    },
}

impl Locked {
    /// Build the handle from the open-time classification.
    #[allow(
        dead_code,
        reason = "encrypted inputs are recognized only in builds that read compound files"
    )]
    pub(crate) fn new(
        container: Container,
        format: Option<FileFormat>,
        encryption: EncryptionKind,
        source: LockedSource,
        options: &OpenOptions,
    ) -> Self {
        Self::new_with_origin(container, format, encryption, source, options, None)
    }

    /// Build a locked handle while retaining the stable path origin of the
    /// open. In-memory routes continue to use [`Self::new`] and leave it
    /// unset.
    pub(crate) fn new_with_origin(
        container: Container,
        format: Option<FileFormat>,
        encryption: EncryptionKind,
        source: LockedSource,
        options: &OpenOptions,
        origin: Option<Arc<std::path::Path>>,
    ) -> Self {
        Self(Arc::new(LockedState {
            container,
            format,
            encryption,
            source,
            options: options.clone(),
            origin,
        }))
    }

    /// A locked handle carrying the same source and a path origin supplied by
    /// the path route after content classification.
    #[cfg(any(
        feature = "doc",
        feature = "xls",
        feature = "docx",
        feature = "ppt",
        feature = "xlsx",
        feature = "pptx",
        feature = "xlsb"
    ))]
    #[allow(
        dead_code,
        reason = "used by the OOXML path-origin route when that feature is enabled"
    )]
    pub(crate) fn with_origin(&self, origin: Option<Arc<std::path::Path>>) -> Self {
        Self(Arc::new(self.0.clone_with_origin(origin)))
    }

    /// The format of the encrypted content.
    ///
    /// Legacy formats keep their identity in clear streams, so an encrypted
    /// DOC, XLS or PPT always reports it. OOXML encryption covers the whole
    /// package: an encrypted OOXML package reports the format its file name
    /// names, and `None` when it was opened from memory or its name does not
    /// name an OOXML format. Unlocking reveals the format either way.
    #[must_use]
    pub fn format(&self) -> Option<FileFormat> {
        self.0.format
    }

    /// The family of [`Self::format`].
    #[must_use]
    pub fn family(&self) -> Option<Family> {
        self.0.format.map(Family::of)
    }

    /// The encryption scheme, classified without a password.
    #[must_use]
    pub fn encryption(&self) -> EncryptionKind {
        self.0.encryption
    }

    /// The error every read and edit of a locked file returns.
    pub(crate) fn read_error(&self) -> Error {
        Error::locked(self.0.format)
    }

    /// The error a `family` view of a locked file returns: the wrong-family
    /// refusal when the format is known and belongs to another family, and
    /// [`Self::read_error`] otherwise.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "rtf",
        feature = "odt",
        feature = "ppt",
        feature = "pptx",
        feature = "odp",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "ods"
    ))]
    pub(crate) fn view_error(&self, family: Family) -> Error {
        match self.family() {
            Some(known) if known != family => Error::unsupported_format(self.0.format),
            _ => self.read_error(),
        }
    }
}

impl LockedState {
    /// Clone the state envelope while sharing positional sources and
    /// preserving the open policy. Owned bytes are cloned only when a caller
    /// needs an independent locked handle.
    #[cfg(any(
        feature = "doc",
        feature = "xls",
        feature = "docx",
        feature = "ppt",
        feature = "xlsx",
        feature = "pptx",
        feature = "xlsb",
        feature = "encryption"
    ))]
    #[allow(
        dead_code,
        reason = "used by the OOXML path-origin transition in feature-enabled builds"
    )]
    fn clone_with_origin(&self, origin: Option<Arc<std::path::Path>>) -> Self {
        Self {
            container: self.container,
            format: self.format,
            encryption: self.encryption,
            source: match &self.source {
                LockedSource::Owned(bytes) => LockedSource::Owned(bytes.clone()),
                #[cfg(any(unix, windows))]
                LockedSource::Positional { source, version } => LockedSource::Positional {
                    source: Arc::clone(source),
                    version: *version,
                },
            },
            options: self.options.clone(),
            origin,
        }
    }
}

#[cfg(all(
    feature = "encryption",
    any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp",
        feature = "pages",
        feature = "keynote",
        feature = "numbers",
        odf_objects
    )
))]
impl Locked {
    /// Consume the handle and open its content with `password`.
    ///
    /// The source the open retained is decrypted once; the input is not
    /// opened again. The unlocked [`File`](crate::File) uses the eager owner,
    /// which keeps the encryption profile, so a later save cannot write the
    /// content in the clear unless asked to. A file whose signature survives
    /// decryption opens as [`File::Signed`](crate::File::Signed).
    ///
    /// # Errors
    ///
    /// Returns [`Code::InvalidArgument`](crate::Code::InvalidArgument) with a
    /// suggestion when the password is wrong,
    /// [`Code::UnsupportedCapability`](crate::Code::UnsupportedCapability)
    /// for a scheme or format this build cannot decrypt (an encrypted XLSB
    /// package, for instance), [`Code::IoError`](crate::Code::IoError) when a
    /// file source changed after open, and the format's typed error when the
    /// decrypted content cannot be read.
    ///
    /// ```no_run
    /// use litchi_ops::{File, Password};
    ///
    /// # fn main() -> litchi_ops::Result<()> {
    /// let file = match litchi_ops::open("secret.docx")? {
    ///     File::Locked(locked) => locked.unlock(Password::new("pa55word".to_owned()))?,
    ///     file => file,
    /// };
    /// println!("{}", file.text()?);
    /// # Ok(())
    /// # }
    /// ```
    pub fn unlock(self, password: crate::Password) -> crate::Result<crate::File> {
        let state = match Arc::try_unwrap(self.0) {
            Ok(state) => state,
            Err(shared) => shared.duplicate(),
        };
        let bytes = state.source.into_bytes(&state.options)?;
        crate::open::unlock_with_origin(
            state.container,
            bytes,
            password.as_str(),
            state.format,
            &state.options,
            state.origin,
        )
    }
}

#[cfg(all(
    feature = "encryption",
    any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp",
        feature = "pages",
        feature = "keynote",
        feature = "numbers",
        odf_objects
    )
))]
impl LockedState {
    /// A copy of a state another clone still shares. Only an in-memory
    /// source is copied; a positional source is shared.
    fn duplicate(&self) -> Self {
        self.clone_with_origin(self.origin.clone())
    }
}

#[cfg(all(
    feature = "encryption",
    any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp",
        feature = "pages",
        feature = "keynote",
        feature = "numbers",
        odf_objects
    )
))]
impl LockedSource {
    /// The complete input, read once from a positional source under the
    /// open's input ceiling and checked against the revision opened.
    fn into_bytes(self, options: &OpenOptions) -> crate::Result<Vec<u8>> {
        match self {
            Self::Owned(bytes) => Ok(bytes),
            #[cfg(any(unix, windows))]
            Self::Positional { source, version } => crate::open::read_whole_source(
                source.as_ref(),
                version,
                crate::open::input_ceiling(options),
            ),
        }
    }
}

impl fmt::Debug for Locked {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Locked")
            .field("format", &self.0.format)
            .field("encryption", &self.0.encryption)
            .finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const fn assert_traits<T: Send + Sync + Clone + fmt::Debug>() {}
    const _: () = {
        assert_traits::<EncryptionKind>();
        assert_traits::<Locked>();
    };
    // One pointer, so `File` stays a tag plus one pointer (record 0003 D2).
    const _: () = assert!(size_of::<Locked>() == size_of::<usize>());

    fn locked(format: Option<FileFormat>, encryption: EncryptionKind) -> Locked {
        Locked::new(
            Container::Ooxml,
            format,
            encryption,
            LockedSource::Owned(vec![1, 2, 3]),
            &OpenOptions::default(),
        )
    }

    #[test]
    fn locked_reports_format_and_kind() {
        let locked = locked(Some(FileFormat::Docx), EncryptionKind::Agile);
        assert_eq!(locked.format(), Some(FileFormat::Docx));
        assert_eq!(locked.family(), Some(Family::Document));
        assert_eq!(locked.encryption(), EncryptionKind::Agile);
        let clone = locked.clone();
        assert!(Arc::ptr_eq(&locked.0, &clone.0), "clone shares the state");
    }

    #[test]
    fn locked_without_a_known_format_reports_none() {
        let locked = locked(None, EncryptionKind::Standard);
        assert_eq!(locked.format(), None);
        assert_eq!(locked.family(), None);
        let error = locked.read_error();
        assert_eq!(error.code(), crate::Code::DocumentProtected);
        assert_eq!(error.format(), None);
    }

    #[test]
    fn locked_debug_is_redacted() {
        let locked = locked(Some(FileFormat::Xls), EncryptionKind::BinaryRc4);
        assert_eq!(
            format!("{locked:?}"),
            "Locked { format: Some(Xls), encryption: BinaryRc4, .. }"
        );
    }
}
