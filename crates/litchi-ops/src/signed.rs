//! The [`Signed`] state of [`File`]: a readable input whose signature must
//! be consumed explicitly before it is edited (ADR 0006 `Signed<T>`; record
//! 0003 §2.6; record 0007 Q4).

use std::fmt;
use std::sync::Arc;

use litchi_core::FileFormat;

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
/// `Signed` adds no way to edit the file it wraps. The consuming transitions
/// that remove or renew the signature, and the refusal of edit verbs on a
/// signed input, are added by the signature policy step (M0a S6b), so an edit
/// can never strip a signature silently.
///
/// `Signed` is a cheap-to-clone handle. Its [`Debug`] output names the inner
/// format and the signature origin only.
///
/// ```no_run
/// use litchi_ops::File;
///
/// # fn main() -> litchi_ops::Result<()> {
/// if let File::Signed(signed) = litchi_ops::open("contract.docx")? {
///     println!("{:?} signed in its {:?} container", signed.format(), signed.origin());
///     println!("{}", signed.file().text()?);
/// }
/// # Ok(())
/// # }
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
    /// `origin`. The openers never wrap a [`File::Locked`] or a nested
    /// [`File::Signed`].
    #[cfg_attr(
        not(any(
            feature = "doc",
            feature = "docx",
            feature = "ppt",
            feature = "pptx",
            feature = "xls",
            feature = "xlsx",
            feature = "xlsb"
        )),
        expect(
            dead_code,
            reason = "only OPC and compound-file formats carry detectable signatures"
        )
    )]
    pub(crate) fn new(file: File, format: FileFormat, origin: SignatureOrigin) -> Self {
        Self(Arc::new(SignedState {
            file,
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

    /// The readable file the signature covers.
    #[must_use]
    pub fn file(&self) -> &File {
        &self.0.file
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
