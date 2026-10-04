//! The flat `File` enum, including not-yet-editable `Locked` and `Signed`
//! states (design section 4; record 0003 §2.1 and §2.6).
//!
//! This module currently holds the pieces of that shape that do not depend on
//! any per-format handle: [`Family`], and the [`Locked`] and [`Signed`]
//! shells. The `File` enum itself, `Locked::unlock`, and the `Signed`
//! consuming transitions arrive in M0a steps S6 and S6b.

use std::fmt;
use std::sync::Arc;

use litchi_core::FileFormat;

/// The broad family of a format, reported as data rather than as an enum
/// layer around the per-format variants (record 0003 §2.1, decision D1).
///
/// ```
/// use litchi_core::FileFormat;
/// use litchi_ops::Family;
///
/// assert_eq!(Family::of(FileFormat::Docx), Family::Document);
/// assert_eq!(Family::of(FileFormat::Odp), Family::Presentation);
/// assert_eq!(Family::of(FileFormat::Xlsb), Family::Workbook);
/// assert_eq!(Family::of(FileFormat::Odg), Family::Other);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Family {
    /// Word-processing documents: DOC, DOCX, RTF, ODT, ODM, OTH, Pages.
    Document,
    /// Presentations: PPT, PPTX, ODP, Keynote.
    Presentation,
    /// Spreadsheets: XLS, XLSX, XLSB, ODS, Numbers.
    Workbook,
    /// Formats outside the three families: drawings, charts, images,
    /// databases and formulas (ODG, ODC, ODI, ODB, ODF formula).
    Other,
}

impl Family {
    /// The family `format` belongs to.
    ///
    /// ```
    /// use litchi_core::FileFormat;
    /// use litchi_ops::Family;
    ///
    /// assert_eq!(Family::of(FileFormat::Rtf), Family::Document);
    /// ```
    #[must_use]
    pub const fn of(format: FileFormat) -> Self {
        match format {
            FileFormat::Doc
            | FileFormat::Docx
            | FileFormat::Rtf
            | FileFormat::Odt
            | FileFormat::Odm
            | FileFormat::Oth
            | FileFormat::Pages => Self::Document,
            FileFormat::Ppt | FileFormat::Pptx | FileFormat::Odp | FileFormat::Keynote => {
                Self::Presentation
            },
            FileFormat::Xls
            | FileFormat::Xlsx
            | FileFormat::Xlsb
            | FileFormat::Ods
            | FileFormat::Numbers => Self::Workbook,
            _ => Self::Other,
        }
    }
}

/// The encryption scheme of a [`Locked`] input, as classified without a
/// password.
///
/// OOXML kinds come from `litchi_crypto::ooxml::inspect` when the
/// `encryption` feature is on; legacy kinds come from the leaf's
/// password-required report. [`Self::Unknown`] means the input was
/// recognized as encrypted (for example from stream presence) but its scheme
/// was not classified.
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

/// An encrypted input: recognized, but neither readable nor editable until it
/// is unlocked with a password (ADR 0006 `Locked<T>`).
///
/// `Locked` is a cheap-to-clone handle. Its [`Debug`] output names the inner
/// format and encryption kind only; it never prints source bytes or key
/// material.
///
/// `Locked` values are produced by `open`. The consuming
/// `Locked::unlock(self, password)` transition is added in M0a step S6.
#[derive(Clone)]
pub struct Locked(Arc<LockedState>);

/// Shared state behind a [`Locked`] handle. Step S6 adds the retained source
/// the `unlock` transition decrypts from.
struct LockedState {
    format: FileFormat,
    encryption: EncryptionKind,
}

impl Locked {
    /// Build the handle from the open-time classification.
    #[cfg_attr(
        not(test),
        expect(
            dead_code,
            reason = "constructed by open() classification in M0a step S6"
        )
    )]
    pub(crate) fn new(format: FileFormat, encryption: EncryptionKind) -> Self {
        Self(Arc::new(LockedState { format, encryption }))
    }

    /// The format of the encrypted content, as detected from the container.
    #[must_use]
    pub fn format(&self) -> FileFormat {
        self.0.format
    }

    /// The family of [`Self::format`].
    #[must_use]
    pub fn family(&self) -> Family {
        Family::of(self.0.format)
    }

    /// The encryption scheme, classified without a password.
    #[must_use]
    pub fn encryption(&self) -> EncryptionKind {
        self.0.encryption
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

/// A signed input: readable, but every edit verb refuses until the signature
/// is consumed explicitly (ADR 0006 `Signed<T>`; record 0007 Q4: every signed
/// input is wrapped).
///
/// `Signed` is a cheap-to-clone handle. Its [`Debug`] output names the inner
/// format and signature origin only; it never prints document content or
/// signature material.
///
/// `Signed` values are produced by `open`. Step S6 makes the wrapped `File`
/// reachable through `Signed::file`, and step S6b adds the consuming
/// `unsign` transition.
#[derive(Clone)]
pub struct Signed(Arc<SignedState>);

/// Shared state behind a [`Signed`] handle. Step S6 replaces `format` with
/// the wrapped `File`, whose format `Signed::format` then reports.
struct SignedState {
    format: FileFormat,
    origin: SignatureOrigin,
}

impl Signed {
    /// Build the handle from the open-time classification.
    #[cfg_attr(
        not(test),
        expect(
            dead_code,
            reason = "constructed by open() classification in M0a step S6"
        )
    )]
    pub(crate) fn new(format: FileFormat, origin: SignatureOrigin) -> Self {
        Self(Arc::new(SignedState { format, origin }))
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
        assert_traits::<Family>();
        assert_traits::<EncryptionKind>();
        assert_traits::<SignatureOrigin>();
        assert_traits::<Locked>();
        assert_traits::<Signed>();
    };
    // One pointer each, so `File` can stay a tag plus one pointer (D2).
    const _: () = assert!(size_of::<Locked>() == size_of::<usize>());
    const _: () = assert!(size_of::<Signed>() == size_of::<usize>());

    #[test]
    fn family_covers_every_detected_format() {
        use FileFormat as F;
        let cases = [
            (F::Doc, Family::Document),
            (F::Docx, Family::Document),
            (F::Rtf, Family::Document),
            (F::Odt, Family::Document),
            (F::Odm, Family::Document),
            (F::Oth, Family::Document),
            (F::Pages, Family::Document),
            (F::Ppt, Family::Presentation),
            (F::Pptx, Family::Presentation),
            (F::Odp, Family::Presentation),
            (F::Keynote, Family::Presentation),
            (F::Xls, Family::Workbook),
            (F::Xlsx, Family::Workbook),
            (F::Xlsb, Family::Workbook),
            (F::Ods, Family::Workbook),
            (F::Numbers, Family::Workbook),
            (F::Odg, Family::Other),
            (F::Odc, Family::Other),
            (F::Odi, Family::Other),
            (F::Odb, Family::Other),
            (F::Odf, Family::Other),
        ];
        for (format, family) in cases {
            assert_eq!(Family::of(format), family, "{format:?}");
        }
    }

    #[test]
    fn locked_reports_format_and_kind() {
        let locked = Locked::new(FileFormat::Docx, EncryptionKind::Agile);
        assert_eq!(locked.format(), FileFormat::Docx);
        assert_eq!(locked.family(), Family::Document);
        assert_eq!(locked.encryption(), EncryptionKind::Agile);
        let clone = locked.clone();
        assert!(Arc::ptr_eq(&locked.0, &clone.0), "clone shares the state");
    }

    #[test]
    fn locked_debug_is_redacted() {
        let locked = Locked::new(FileFormat::Xls, EncryptionKind::BinaryRc4);
        assert_eq!(
            format!("{locked:?}"),
            "Locked { format: Xls, encryption: BinaryRc4, .. }"
        );
    }

    #[test]
    fn signed_reports_format_and_origin() {
        let signed = Signed::new(FileFormat::Pptx, SignatureOrigin::Opc);
        assert_eq!(signed.format(), FileFormat::Pptx);
        assert_eq!(signed.family(), Family::Presentation);
        assert_eq!(signed.origin(), SignatureOrigin::Opc);
        let clone = signed.clone();
        assert!(Arc::ptr_eq(&signed.0, &clone.0), "clone shares the state");
    }

    #[test]
    fn signed_debug_is_redacted() {
        let signed = Signed::new(FileFormat::Doc, SignatureOrigin::Cfb);
        assert_eq!(
            format!("{signed:?}"),
            "Signed { format: Doc, origin: Cfb, .. }"
        );
    }
}
