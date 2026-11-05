//! The [`Signed`] state of [`File`]: a readable input whose signature must
//! be consumed explicitly before it is edited (ADR 0006 `Signed<T>`; record
//! 0003 §2.6; record 0007 Q4; record 0013, the signature policy).

use std::fmt;
use std::sync::Arc;

use litchi_core::FileFormat;

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb"
))]
use crate::format::signature::Signature;
use crate::{Family, File};

/// Where a [`Signed`] input's signature was found. Detection is
/// metadata-only; nothing is verified at open.
///
/// ```
/// use litchi_ops::SignatureOrigin;
///
/// assert_ne!(SignatureOrigin::Opc, SignatureOrigin::Cfb);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SignatureOrigin {
    /// OPC digital-signature origin part and its signature parts.
    Opc,
    /// Compound-file `_signatures` or `_xmlsignatures` streams.
    Cfb,
}

/// A signed input. Every signed input opens as `Signed` (record 0007 Q4).
///
/// Reads go to the [`File`] it wraps: [`File::format`], [`File::text`] and
/// the other reads of a [`File::Signed`] answer for the wrapped file, and
/// [`Signed::file`] returns it for format-specific reads. Opening verifies
/// nothing; the signature is only detected, from package metadata.
///
/// Changed editing must consume the signature explicitly (ADR 0006), so that
/// no edit can strip a signature silently:
///
/// - An empty or exact no-op edit may commit and preserves the signed
///   publication byte-identically. That includes the wrapped file returned by
///   [`Signed::file`] and the outer `File::Signed` facade.
/// - A changed edit while the file is signed refuses with
///   [`Code::DocumentProtected`](crate::Code::DocumentProtected) and a
///   suggestion naming [`Signed::unsign`]. The wrapped file follows the same
///   rule: [`Signed::file`] can prove an exact no-op but cannot publish a
///   changed result.
/// - [`Signed::unsign`] consumes the signature and returns the file. Its
///   first write removes the signature parts, through the format's own
///   explicit removal, and publishes the document without them.
/// - Re-signing with caller keys (`resign`) is deferred: see
///   [`Signed::unsign`].
///
/// `Signed` is a cheap-to-clone handle. Its [`Debug`] output names the inner
/// format and the signature origin only.
///
/// ```no_run
/// use litchi_ops::File;
///
/// # #[cfg(feature = "docx")]
/// # fn main() -> litchi_ops::Result<()> {
/// if let File::Signed(signed) = litchi_ops::open("contract.docx")? {
///     println!("{:?} signed in its {:?} container", signed.format(), signed.origin());
///     println!("{}", signed.file().text()?);
///     // Editing requires removing the signature explicitly.
///     let unsigned: File = signed.unsign();
///     assert!(!unsigned.is_signed());
/// }
/// # Ok(())
/// # }
/// # #[cfg(not(feature = "docx"))]
/// # fn main() {}
/// ```
#[derive(Clone)]
pub struct Signed(Arc<SignedState>);

/// Shared state behind a [`Signed`] handle. The `Arc` breaks the recursion
/// between [`File`] and the file a `Signed` wraps.
struct SignedState {
    file: File,
    format: FileFormat,
    origin: SignatureOrigin,
}

impl Signed {
    /// Wrap the readable `file` of `format` whose signature was found in
    /// `origin`, marking its handle signed so that its editor builders
    /// refuse. The openers never wrap a [`File::Locked`] or a nested
    /// [`File::Signed`].
    /// Only OPC and compound-file formats carry detectable signatures.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb"
    ))]
    pub(crate) fn new(file: File, format: FileFormat, origin: SignatureOrigin) -> Self {
        Self(Arc::new(SignedState {
            file: file.with_signature(Signature::Signed),
            format,
            origin,
        }))
    }

    /// The format of the signed content.
    #[must_use]
    pub fn format(&self) -> FileFormat {
        self.0.format
    }

    /// The family of [`Self::format`].
    #[must_use]
    pub fn family(&self) -> Family {
        Family::of(self.0.format)
    }

    /// Where the signature was found.
    #[must_use]
    pub fn origin(&self) -> SignatureOrigin {
        self.0.origin
    }

    /// The readable file the signature covers. It reads like any file of
    /// its format; exact no-op edits preserve its signed bytes, while changed
    /// edits refuse until [`Self::unsign`] consumes the signature.
    #[must_use]
    pub fn file(&self) -> &File {
        &self.0.file
    }

    /// Consume the signature explicitly (ADR 0006) and return the file.
    ///
    /// Nothing is read or rewritten here. The returned file reads the same
    /// content, and [`File::is_signed`] is `false` for it. Its first write
    /// removes the signature through the format's own explicit removal (the
    /// OOXML owner's `unsign`, or the compound-file signature editor's
    /// `clear`) and publishes the document without any signature part; an
    /// edit is then allowed although it invalidates what was signed. Other
    /// clones of this `Signed` stay signed: the authorization belongs to the
    /// returned file only.
    ///
    /// Re-signing with caller keys (`resign`) is deferred until the signing
    /// providers have an explicit, credential-safe input in the unified layer
    /// (record 0013); the format crates keep their own signing APIs.
    ///
    /// ```no_run
    /// use litchi_ops::File;
    ///
    /// # #[cfg(feature = "docx")]
    /// # fn main() -> litchi_ops::Result<()> {
    /// let file = match litchi_ops::open("contract.docx")? {
    ///     File::Signed(signed) => signed.unsign(),
    ///     file => file,
    /// };
    /// assert!(!file.is_signed());
    /// # Ok(())
    /// # }
    /// # #[cfg(not(feature = "docx"))]
    /// # fn main() {}
    /// ```
    #[must_use = "the unsigned file is the only handle the signature removal authorizes"]
    pub fn unsign(self) -> File {
        #[cfg(any(
            feature = "doc",
            feature = "docx",
            feature = "ppt",
            feature = "pptx",
            feature = "xls",
            feature = "xlsx",
            feature = "xlsb"
        ))]
        {
            self.0.file.with_signature(Signature::RemovalAuthorized)
        }
        #[cfg(not(any(
            feature = "doc",
            feature = "docx",
            feature = "ppt",
            feature = "pptx",
            feature = "xls",
            feature = "xlsx",
            feature = "xlsb"
        )))]
        {
            self.0.file.clone()
        }
    }
}

impl fmt::Debug for Signed {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Signed")
            .field("format", &self.0.format)
            .field("origin", &self.0.origin)
            .finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const fn assert_traits<T: Send + Sync + Clone + fmt::Debug>() {}
    const _: () = {
        assert_traits::<SignatureOrigin>();
        assert_traits::<Signed>();
    };
    // One pointer, so `File` stays a tag plus one pointer (record 0003 D2).
    const _: () = assert!(size_of::<Signed>() == size_of::<usize>());
}
