//! Protected-container-aware access to checked CFB overlay publication.
//!
//! This is a narrow physical publication primitive for format owners that
//! have already validated an equal-length semantic edit. It does not replace
//! semantic candidate validation and it deliberately does not fall back to a
//! topology-changing render.

use crate::protection::reject_protected_shared_container;
use litchi_cfb::{
    ComposedOverlaySource, OverlayError, OverlayLimits, SameLengthStreamOverlay,
    SameLengthStreamSplice, SharedOleFile, SharedOleFileLimits, StreamSpliceLimits,
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
    cfb: Arc<SharedOleFile>,
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

    /// Opens exact immutable bytes without erasing their owned provenance.
    ///
    /// The common protected-container guard is identical to [`Self::open`].
    /// Only the low-level CFB owner may use the provenance to specialize a
    /// later publication. Composed views retain their normal complete
    /// fingerprint fence; atomic saves may omit only the two redundant outer
    /// fences while keeping the full source/target emission hashes and all
    /// flush, fsync, rename, and parent-sync durability steps.
    pub fn open_owned(source: Arc<[u8]>, version: SourceVersion) -> Result<Self, OverlayError> {
        let cfb = SharedOleFile::open_owned(source, version)?;
        reject_protected_shared_container(&cfb, "source-backed stream overlay publication")?;
        Ok(Self { cfb: Arc::new(cfb) })
    }

    /// Opens under a caller-selected positional CFB ingress ceiling, then
    /// applies the common signed/encrypted/DRM mutation guard.
    pub fn open_with_limits(
        source: Arc<dyn ReadAt>,
        limits: SharedOleFileLimits,
    ) -> Result<Self, OverlayError> {
        let cfb = SharedOleFile::open_with_limits(source, limits)?;
        reject_protected_shared_container(&cfb, "source-backed stream overlay publication")?;
        Ok(Self { cfb: Arc::new(cfb) })
    }

    /// Returns the validated positional CFB owner for repeated semantic reads.
    ///
    /// The returned handle retains the parsed directory/FAT index and shares
    /// the same source-version and lazy mini-stream state as this publisher.
    /// Callers must still perform their operation-level source-version and
    /// fingerprint checks before using it.
    pub fn shared(&self) -> Arc<SharedOleFile> {
        Arc::clone(&self.cfb)
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

    /// Builds a fully reopened, reusable plan for bounded same-length ranges.
    ///
    /// Each splice is checked against the source before a plan is returned;
    /// this avoids staging a complete replacement stream when a semantic
    /// owner already has exact source-relative ranges to publish. The common
    /// source version, protected-container checks, full artifact fingerprints,
    /// and candidate CFB reopen remain unchanged from [`Self::plan`].
    pub fn plan_splices(
        &self,
        splices: Vec<SameLengthStreamSplice>,
        limits: StreamSpliceLimits,
    ) -> Result<ValidatedOverlayPlan, OverlayError> {
        self.cfb.plan_same_length_stream_splices(splices, limits)
    }

    /// Builds a splice plan whose effective target is also checked by a
    /// format-owner validator inside the complete CFB fingerprint fence.
    ///
    /// The owner receives only the lazy composed target and its native error
    /// type is preserved. Exact byte no-ops skip the callback and return no
    /// owner result. Protected-container refusal remains an ingress property
    /// of this publisher and publication remains owned by the returned plan.
    pub fn plan_splices_with_owner<T, E, F>(
        &self,
        splices: Vec<SameLengthStreamSplice>,
        limits: StreamSpliceLimits,
        validate_owner: F,
    ) -> Result<(ValidatedOverlayPlan, Option<T>), E>
    where
        F: FnOnce(&ComposedOverlaySource) -> Result<T, E>,
        E: From<OverlayError>,
    {
        self.cfb
            .plan_same_length_stream_splices_with_owner(splices, limits, validate_owner)
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
    fn wrapper_publishes_only_the_validated_equal_length_splice() {
        let bytes = write_cfb(|writer| {
            writer.create_stream(&["Target"], b"before").unwrap();
            writer.create_stream(&["Opaque"], b"unchanged").unwrap();
        });
        let publisher =
            SourceBackedOverlayPublisher::open(Arc::new(OwnedSource::new(bytes))).unwrap();
        let plan = publisher
            .plan_splices(
                vec![SameLengthStreamSplice::new(
                    vec!["Target".to_string()],
                    2,
                    Arc::from(b"fo".as_slice()),
                    Arc::from(b"XX".as_slice()),
                )],
                StreamSpliceLimits::default(),
            )
            .unwrap();
        let mut output = Vec::new();
        plan.write_to(&mut output).unwrap();
        let mut ole = OleFile::open(Cursor::new(output)).unwrap();
        assert_eq!(ole.open_stream(&["Target"]).unwrap(), b"beXXre");
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
                    Arc::new(OwnedSource::new(bytes.clone())),
                    SharedOleFileLimits::default(),
                ),
                Err(OverlayError::Ole(litchi_cfb::OleError::InvalidFormat(_)))
            ));
            assert!(matches!(
                SourceBackedOverlayPublisher::open_owned(
                    Arc::from(bytes),
                    SourceVersion::new(0x0181, 0),
                ),
                Err(OverlayError::Ole(litchi_cfb::OleError::InvalidFormat(_)))
            ));
        }
    }
}
