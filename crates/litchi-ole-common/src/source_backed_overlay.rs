//! Protected-container-aware access to checked CFB overlay publication.
//!
//! This is a narrow physical publication primitive for format owners that
//! have already validated an equal-length semantic edit. It does not replace
//! semantic candidate validation and it deliberately does not fall back to a
//! topology-changing render.

use crate::protection::reject_protected_shared_container;
use litchi_cfb::{
    OverlayError, OverlayLimits, SameLengthStreamOverlay, SharedOleFile, SharedOleFileLimits,
    ValidatedOverlayPlan,
};
use litchi_core::{ReadAt, SourceVersion};
use std::sync::Arc;

/// A validated positional OLE source eligible for unchanged-topology stream
/// overlays.
///
/// Opening applies the same signing, encryption, and DRM refusal as the common
/// object editor, including markers represented by empty storages.
pub struct SourceBackedOverlayPublisher {
    cfb: SharedOleFile,
}

impl std::fmt::Debug for SourceBackedOverlayPublisher {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SourceBackedOverlayPublisher")
            .field("file_size", &self.cfb.file_size())
            .finish_non_exhaustive()
    }
}

impl SourceBackedOverlayPublisher {
    /// Opens and fully validates a bounded positional CFB source, then applies
    /// the common signed/encrypted/DRM mutation guard.
    pub fn open(source: Arc<dyn ReadAt>) -> Result<Self, OverlayError> {
        Self::open_with_limits(source, SharedOleFileLimits::default())
    }

    /// Opens under a caller-selected positional CFB ingress ceiling, then
    /// applies the common signed/encrypted/DRM mutation guard.
    pub fn open_with_limits(
        source: Arc<dyn ReadAt>,
        limits: SharedOleFileLimits,
    ) -> Result<Self, OverlayError> {
        let cfb = SharedOleFile::open_with_limits(source, limits)?;
        reject_protected_shared_container(&cfb, "source-backed stream overlay publication")?;
        Ok(Self { cfb })
    }

    /// Exact source identity/revision captured during open.
    pub fn source_version(&self) -> Result<SourceVersion, OverlayError> {
        self.cfb.source_version().map_err(OverlayError::from)
    }

    /// Builds a fully reopened, reusable same-length publication plan.
    ///
    /// This method never invokes the existing full-render fallback. Callers
    /// must explicitly choose that path for length or topology changes.
    pub fn plan(
        &self,
        overlays: Vec<SameLengthStreamOverlay>,
        limits: OverlayLimits,
    ) -> Result<ValidatedOverlayPlan, OverlayError> {
        self.cfb.plan_same_length_stream_overlays(overlays, limits)
    }
}

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used, reason = "test assertions panic by design")]

    use super::*;
    use litchi_cfb::{OleFile, OleWriter};
    use litchi_core::OwnedSource;
    use std::io::Cursor;

    fn write_cfb(build: impl FnOnce(&mut OleWriter)) -> Vec<u8> {
        let mut writer = OleWriter::new();
        build(&mut writer);
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn wrapper_publishes_only_the_validated_equal_length_stream() {
        let bytes = write_cfb(|writer| {
            writer.create_stream(&["Target"], b"before").unwrap();
            writer.create_stream(&["Opaque"], b"unchanged").unwrap();
        });
        let publisher =
            SourceBackedOverlayPublisher::open(Arc::new(OwnedSource::new(bytes))).unwrap();
        let plan = publisher
            .plan(
                vec![SameLengthStreamOverlay::new(
                    vec!["Target".to_string()],
                    Arc::from(b"change".as_slice()),
                )],
                OverlayLimits::default(),
            )
            .unwrap();
        let mut output = Vec::new();
        plan.write_to(&mut output).unwrap();
        let mut ole = OleFile::open(Cursor::new(output)).unwrap();
        assert_eq!(ole.open_stream(&["Target"]).unwrap(), b"change");
        assert_eq!(ole.open_stream(&["Opaque"]).unwrap(), b"unchanged");
    }

    #[test]
    fn wrapper_retains_every_existing_protected_component_refusal() {
        let markers = [
            "_xmlsignatures",
            "_signatures",
            "DigitalSignature",
            "\u{0005}DigitalSignature",
            "\u{0006}DataSpaces",
            "\u{0006}DataSpaceInfo",
            "\u{0006}TransformInfo",
            "\u{0006}Primary",
            "\u{0009}DRMContent",
            "\u{0009}DRMViewerContent",
            "EncryptedPackage",
            "EncryptionInfo",
        ];
        for marker in markers {
            let bytes = write_cfb(|writer| {
                writer.create_storage(&[marker]).unwrap();
                writer.create_stream(&["Ordinary"], b"bytes").unwrap();
            });
            assert!(matches!(
                SourceBackedOverlayPublisher::open_with_limits(
                    Arc::new(OwnedSource::new(bytes)),
                    SharedOleFileLimits::default(),
                ),
                Err(OverlayError::Ole(litchi_cfb::OleError::InvalidFormat(_)))
            ));
        }
    }
}
