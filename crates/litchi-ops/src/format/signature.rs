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

    /// The handle of a committed state with its signature disposition and
    /// path origin, built in one step (record 0286): adopting a commit
    /// through [`Self::new`], [`Self::with_origin`] and
    /// [`Self::with_signature`] allocates two handles only to drop them.
    #[cfg(feature = "docx")]
    pub(crate) fn adopted(state: S, signature: Signature, origin: Option<Arc<Path>>) -> Arc<Self> {
        Self::from_parts_with_origin(signature, Arc::new(state), origin)
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

    /// A handle over the same state with another disposition.
    pub(crate) fn with_signature(&self, signature: Signature) -> Arc<Self> {
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
