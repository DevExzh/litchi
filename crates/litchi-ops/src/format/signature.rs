//! The signature disposition of the tier 1 and tier 2 handles (ADR 0006;
//! unified-ops record 0013).
//!
//! An open that finds a signature wraps the file in [`crate::Signed`] and
//! marks its handle [`Signature::Signed`]: reads work, and every editor
//! builder refuses with `document_protected` and a suggestion naming
//! [`crate::Signed::unsign`]. That consuming transition is the only way out.
//! It marks the same file [`Signature::RemovalAuthorized`] without reading or
//! rewriting anything. The first editor built from such a handle works on a
//! copy of the input whose signature the leaf removed explicitly
//! (`unsign` on the OOXML leaf owner, `litchi_sign::cfb::Editor::clear` on a
//! compound file), so every write from that editor publishes without the
//! signature parts. Nothing else in the unified layer removes a signature.

use std::ops::Deref;
use std::path::Path;
use std::sync::Arc;

use litchi_core::FileFormat;

use crate::error::Error;
#[cfg(any(feature = "doc", feature = "ppt", feature = "xls", feature = "xlsb"))]
use crate::error::Result;

/// What a handle may do with the signature of the input it was opened from.
///
/// The transitions, each a method below and the only way a handle changes
/// its disposition ([`Opened`] has no general setter):
///
/// | from                | `found`  | `authorize_removal` | `after_commit` |
/// |---------------------|----------|---------------------|----------------|
/// | `Unsigned`          | `Signed` | `Unsigned`          | `Unsigned`     |
/// | `Signed`            | `Signed` | `RemovalAuthorized` | `Signed`       |
/// | `RemovalAuthorized` | `Signed` | `RemovalAuthorized` | `Unsigned`     |
///
/// ```text
/// Unsigned --found--> Signed --authorize_removal--> RemovalAuthorized --after_commit--> Unsigned
/// ```
///
/// No transition leads from `Signed` to `Unsigned`: a signature is only ever
/// dropped by a committed write of a handle [`crate::Signed::unsign`]
/// returned.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Signature {
    /// The opener found no signature.
    Unsigned,
    /// The input is signed: this is the file a [`crate::Signed`] wraps. Reads
    /// work; every editor builder refuses.
    Signed,
    /// [`crate::Signed::unsign`] consumed the signature. Editors work on a
    /// copy whose signature the leaf removed explicitly.
    RemovalAuthorized,
}

impl Signature {
    /// The opener found a signature in the container.
    fn found(self) -> Self {
        match self {
            Self::Unsigned | Self::Signed | Self::RemovalAuthorized => Self::Signed,
        }
    }

    /// [`crate::Signed::unsign`] consumed the signed file. A handle that
    /// carries no signature has none to remove and stays as it is.
    fn authorize_removal(self) -> Self {
        match self {
            Self::Signed | Self::RemovalAuthorized => Self::RemovalAuthorized,
            Self::Unsigned => Self::Unsigned,
        }
    }

    /// A write was committed. The editor of a removal-authorized handle works
    /// on a copy without the signature, so the committed state is ordinary
    /// unsigned state. A signed handle only commits exact no-ops, which keep
    /// the signature.
    #[cfg(any(
        feature = "docx",
        feature = "ppt",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb"
    ))]
    fn after_commit(self) -> Self {
        match self {
            Self::Unsigned | Self::RemovalAuthorized => Self::Unsigned,
            Self::Signed => Self::Signed,
        }
    }
}

/// One handle over the state of one open: that state, which every clone and
/// every signature disposition of the open shares, and this handle's
/// disposition.
///
/// [`crate::Signed::unsign`] re-marks a file without copying its state: the
/// unsigned handle shares the owner, the caches and the retained source of
/// the signed one. [`Deref`] reaches the state.
pub(crate) struct Opened<S> {
    signature: Signature,
    state: Arc<S>,
    /// Stable path origin, when this state came from a path opener. This is
    /// handle metadata rather than format state, so signature transitions and
    /// immutable owner replacements can carry it without duplicating fields
    /// in every format-specific state.
    origin: Option<Arc<Path>>,
}

impl<S> Opened<S> {
    /// The handle of a fresh open, which found no signature yet: the opener
    /// re-marks it when the container turns out to be signed.
    pub(crate) fn new(state: S) -> Arc<Self> {
        Self::from_parts_with_origin(Signature::Unsigned, Arc::new(state), None)
    }

    /// The handle of the `state` a write of this handle committed, with this
    /// handle's path origin and the disposition a commit leaves
    /// ([`Signature::after_commit`]), built in one step: no intermediate
    /// handle is allocated only to be dropped.
    #[cfg(any(
        feature = "docx",
        feature = "ppt",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb"
    ))]
    pub(crate) fn committed(&self, state: S) -> Arc<Self> {
        Self::from_parts_with_origin(
            self.signature.after_commit(),
            Arc::new(state),
            self.origin.clone(),
        )
    }

    /// The handle of a `state` that revises this one without publishing a
    /// write of its own: it keeps this handle's disposition and path origin.
    #[cfg(any(feature = "pptx", feature = "xlsb"))]
    pub(crate) fn revised(&self, state: S) -> Arc<Self> {
        Self::from_parts_with_origin(self.signature, Arc::new(state), self.origin.clone())
    }

    fn from_parts_with_origin(
        signature: Signature,
        state: Arc<S>,
        origin: Option<Arc<Path>>,
    ) -> Arc<Self> {
        Arc::new(Self {
            signature,
            state,
            origin,
        })
    }

    /// This handle's signature disposition.
    pub(crate) fn signature(&self) -> Signature {
        self.signature
    }

    /// A handle over the same state, marked signed: the opener found a
    /// signature in the container.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb"
    ))]
    pub(crate) fn signed(&self) -> Arc<Self> {
        self.with_signature(self.signature.found())
    }

    /// A handle over the same state whose signature
    /// [`crate::Signed::unsign`] authorized the first committed write to
    /// remove.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb"
    ))]
    pub(crate) fn removal_authorized(&self) -> Arc<Self> {
        self.with_signature(self.signature.authorize_removal())
    }

    /// A handle over the same state with another disposition. Private: the
    /// transitions of [`Signature`] are the only dispositions a handle takes.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb"
    ))]
    fn with_signature(&self, signature: Signature) -> Arc<Self> {
        Self::from_parts_with_origin(signature, Arc::clone(&self.state), self.origin.clone())
    }

    /// A handle over the same state with a stable path origin. The origin is
    /// deliberately absent from `Debug`; it is persistence metadata, not a
    /// document value.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "xlsx",
        feature = "pptx",
        feature = "xlsb",
        feature = "xls"
    ))]
    pub(crate) fn with_origin(&self, origin: Option<Arc<Path>>) -> Arc<Self> {
        Self::from_parts_with_origin(self.signature, Arc::clone(&self.state), origin)
    }

    /// The stable path origin retained by this open, if any.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "xlsx",
        feature = "pptx",
        feature = "xlsb",
        feature = "xls"
    ))]
    pub(crate) fn origin(&self) -> Option<Arc<Path>> {
        self.origin.clone()
    }
}

impl<S> Deref for Opened<S> {
    type Target = S;

    fn deref(&self) -> &S {
        &self.state
    }
}

/// The refusal of an editor builder on a handle that is still signed.
pub(crate) fn refusal(format: FileFormat) -> Error {
    Error::signed(format)
}

/// The exact bytes a source-backed OPC package was opened from.
///
/// An editor that removes the signature starts from an owning package opened
/// over these bytes (record 0009 D4: promote through the exact source), so
/// every ZIP member the edit and the signature removal do not touch is
/// published byte for byte.
#[cfg(feature = "xlsb")]
pub(crate) fn exact_source(package: &litchi_opc::SourceBackedPackage) -> Result<Vec<u8>> {
    let artifact = package.source_artifact();
    let mut bytes = Vec::new();
    reserve(
        &mut bytes,
        artifact.len(),
        "the exact OPC source of an unsigned copy",
    )?;
    artifact.write_to(&mut bytes)?;
    Ok(bytes)
}

/// `bytes` with every legacy (`_signatures`) and XML (`_xmlsignatures`)
/// signature removed by the compound-file signature leaf.
#[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
pub(crate) fn unsign_compound_file(
    bytes: Vec<u8>,
    format: litchi_sign::cfb::Format,
) -> Result<Vec<u8>> {
    let mut editor =
        litchi_sign::cfb::Editor::open(bytes, format, litchi_sign::Limits::standard())?;
    editor.clear()?;
    Ok(editor.finish()?)
}

/// Reserve exactly `length` bytes and return the length, or fail with a
/// typed limit or allocation error.
#[cfg(feature = "xlsb")]
fn reserve(bytes: &mut Vec<u8>, length: u64, resource: &'static str) -> Result<usize> {
    let Ok(length) = usize::try_from(length) else {
        return Err(
            litchi_core::Error::ResourceLimit(litchi_core::ResourceLimit {
                resource: litchi_core::Resource::InputBytes,
                observed: length,
                limit: u64::try_from(usize::MAX).unwrap_or(u64::MAX),
                scope: Arc::from(resource),
            })
            .into(),
        );
    };
    bytes
        .try_reserve_exact(length)
        .map_err(|source| Error::from(litchi_core::Error::Allocation { resource, source }))?;
    Ok(length)
}
