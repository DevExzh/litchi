//! One classified immutable OOXML package snapshot, prepared once.
//!
//! A detection coordinator may need the neutral format classification *and* the
//! validated package that produced it. Without a handoff, the coordinator
//! throws the package away and the format owner indexes the same container a
//! second time.
//!
//! [`PreparedOoxmlSource`] is the OOXML analogue of the accepted opaque
//! prepared-source shape recorded in ADR 0010: it owns exactly one classified
//! immutable package snapshot and is consumed by exactly one format adapter.
//! It is deliberately opaque — no archive reader, physical member identity,
//! cache lock, or execution handle appears on its public surface. The only
//! thing that leaves it is the [`SourceBackedPackage`] the OOXML owners already
//! adopt, and that exit re-checks the source revision first.
//!
//! The handle carries the *validated main-part content type* rather than a
//! neutral format enum. Mapping content types to a neutral classification
//! belongs to the coordinator (ADR 0010), so this crate stores the evidence and
//! leaves the interpretation alone.

use litchi_core::SourceVersion;

use crate::error::{OpcError, Result};
use crate::source_backed::SourceBackedPackage;

/// An already-indexed, already-classified OOXML package awaiting its adopter.
///
/// Construct one with [`PreparedOoxmlSource::new`] after classifying a package,
/// then hand it to a single format owner through
/// [`PreparedOoxmlSource::into_package`]. The package is *not* re-indexed.
///
/// # Source-version contract
///
/// The revision observed when the handle was prepared is retained. Every
/// accessor that touches the package re-reads the live source revision and
/// refuses a source that changed in the meantime, so a handle can never adopt
/// bytes other than the ones it classified.
pub struct PreparedOoxmlSource {
    package: SourceBackedPackage,
    main_part_content_type: String,
    prepared_version: SourceVersion,
}

impl std::fmt::Debug for PreparedOoxmlSource {
    /// Report the classification evidence without exposing package internals.
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("PreparedOoxmlSource")
            .field("main_part_content_type", &self.main_part_content_type)
            .field("prepared_version", &self.prepared_version)
            .finish_non_exhaustive()
    }
}

impl PreparedOoxmlSource {
    /// Retain a classified package together with the main-part content type
    /// that classified it.
    ///
    /// The content type is validated against the package's own catalog, so a
    /// coordinator cannot label a package with a content type the package does
    /// not declare. No archive work is repeated: the check walks the already
    /// parsed in-memory catalog.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::InvalidContentType`] when `main_part_content_type`
    /// is blank or is absent from the package catalog, and propagates the
    /// package's execution-policy and source-freshness failures unchanged.
    pub fn new(package: SourceBackedPackage, main_part_content_type: &str) -> Result<Self> {
        package.check_execution()?;
        let trimmed = main_part_content_type.trim();
        if trimmed.is_empty() {
            return Err(OpcError::InvalidContentType {
                value: main_part_content_type.to_owned(),
                reason: "a prepared OOXML source requires a main-part content type".to_owned(),
            });
        }
        let declared = package
            .iter_parts()
            .any(|part| part.content_type().eq_ignore_ascii_case(trimmed));
        if !declared {
            return Err(OpcError::InvalidContentType {
                value: main_part_content_type.to_owned(),
                reason: "the package catalog does not declare this main-part content type"
                    .to_owned(),
            });
        }
        let prepared_version = package.source_version()?;
        Ok(Self {
            package,
            main_part_content_type: trimmed.to_owned(),
            prepared_version,
        })
    }

    /// The validated main-part content type that classified this package.
    #[must_use]
    pub fn main_part_content_type(&self) -> &str {
        &self.main_part_content_type
    }

    /// The source revision captured when this handle was prepared.
    ///
    /// This accessor performs no I/O; use [`Self::source_version`] to compare
    /// it against the live source.
    #[must_use]
    pub const fn prepared_source_version(&self) -> SourceVersion {
        self.prepared_version
    }

    /// The live source revision, after verifying the source is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::SourceChanged`] when the positional source no longer
    /// matches the snapshot this handle classified.
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.ensure_current()?;
        Ok(self.prepared_version)
    }

    /// Refuse a source that changed after preparation.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::SourceChanged`] when the live revision differs from
    /// the prepared one, and propagates the package's own freshness failure.
    pub fn ensure_current(&self) -> Result<()> {
        let observed = self.package.source_version()?;
        if observed == self.prepared_version {
            Ok(())
        } else {
            Err(OpcError::SourceChanged {
                expected: self.prepared_version,
                actual: observed,
            })
        }
    }

    /// Hand the classified package to its single format adapter.
    ///
    /// The package is returned exactly as it was indexed — the adopter reuses
    /// the existing catalog instead of building a second one.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::SourceChanged`] when the source changed between
    /// preparation and adoption, and propagates the package's execution-policy
    /// failure.
    pub fn into_package(self) -> Result<SourceBackedPackage> {
        self.ensure_current()?;
        self.package.check_execution()?;
        Ok(self.package)
    }
}

#[cfg(test)]
mod tests {
    use std::io;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, Ordering};

    use litchi_core::{ReadAt, SourceVersion};

    use super::PreparedOoxmlSource;
    use crate::error::OpcError;
    use crate::source_backed::SourceBackedPackage;

    const WORD_MAIN: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

    fn docx_bytes() -> Vec<u8> {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ooxml/docx/comment.docx"),
        )
        .expect("fixture")
    }

    /// A positional source whose reported revision advances on demand, so a
    /// test can simulate a source rewritten between preparation and adoption.
    struct RevisionSwitchSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
    }

    impl RevisionSwitchSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                revision: AtomicU64::new(0),
            }
        }

        fn advance(&self) {
            self.revision.fetch_add(1, Ordering::Relaxed);
        }
    }

    impl ReadAt for RevisionSwitchSource {
        fn len(&self) -> io::Result<u64> {
            u64::try_from(self.bytes.len()).map_err(|_err| io::Error::other("length"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            let Some(tail) = self.bytes.get(start..) else {
                return Ok(0);
            };
            let count = tail.len().min(output.len());
            output[..count].copy_from_slice(&tail[..count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                0x5052_4550,
                self.revision.load(Ordering::Relaxed),
            ))
        }
    }

    fn open_switchable() -> (Arc<RevisionSwitchSource>, SourceBackedPackage) {
        let source = Arc::new(RevisionSwitchSource::new(docx_bytes()));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let package = SourceBackedPackage::from_read_at(read_at).expect("open");
        (source, package)
    }

    #[test]
    fn prepared_source_retains_the_validated_main_part_content_type() {
        let package = SourceBackedPackage::from_vec(docx_bytes()).expect("open");
        let prepared = PreparedOoxmlSource::new(package, WORD_MAIN).expect("prepare");
        assert_eq!(prepared.main_part_content_type(), WORD_MAIN);
        let package = prepared.into_package().expect("adopt");
        assert!(package.iter_parts().count() > 0);
    }

    #[test]
    fn prepared_source_matches_the_declared_content_type_case_insensitively() {
        let package = SourceBackedPackage::from_vec(docx_bytes()).expect("open");
        let prepared =
            PreparedOoxmlSource::new(package, &WORD_MAIN.to_uppercase()).expect("prepare");
        assert!(
            prepared
                .main_part_content_type()
                .eq_ignore_ascii_case(WORD_MAIN)
        );
    }

    #[test]
    fn prepared_source_refuses_a_content_type_the_package_does_not_declare() {
        let package = SourceBackedPackage::from_vec(docx_bytes()).expect("open");
        let Err(error) = PreparedOoxmlSource::new(package, "application/vnd.example.absent+xml")
        else {
            panic!("an undeclared main-part content type must be refused");
        };
        assert!(
            matches!(error, OpcError::InvalidContentType { .. }),
            "{error:?}"
        );
    }

    #[test]
    fn prepared_source_refuses_a_blank_content_type() {
        let package = SourceBackedPackage::from_vec(docx_bytes()).expect("open");
        let Err(error) = PreparedOoxmlSource::new(package, "   ") else {
            panic!("a blank main-part content type must be refused");
        };
        assert!(
            matches!(error, OpcError::InvalidContentType { .. }),
            "{error:?}"
        );
    }

    #[test]
    fn prepared_source_refuses_a_changed_source_before_reporting_a_version() {
        let (source, package) = open_switchable();
        let prepared = PreparedOoxmlSource::new(package, WORD_MAIN).expect("prepare");
        assert_eq!(prepared.prepared_source_version().revision(), 0);

        source.advance();

        let error = prepared
            .ensure_current()
            .expect_err("a changed source must be refused");
        assert!(matches!(error, OpcError::SourceChanged { .. }), "{error:?}");
        let error = prepared
            .source_version()
            .expect_err("a changed source must be refused");
        assert!(matches!(error, OpcError::SourceChanged { .. }), "{error:?}");
    }

    #[test]
    fn prepared_source_refuses_adoption_after_the_source_revision_changes() {
        let (source, package) = open_switchable();
        let prepared = PreparedOoxmlSource::new(package, WORD_MAIN).expect("prepare");
        source.advance();
        let Err(error) = prepared.into_package() else {
            panic!("adoption must refuse a source that changed after preparation");
        };
        assert!(matches!(error, OpcError::SourceChanged { .. }), "{error:?}");
    }

    #[test]
    fn prepared_source_reports_the_live_revision_while_the_source_is_stable() {
        let (_source, package) = open_switchable();
        let prepared = PreparedOoxmlSource::new(package, WORD_MAIN).expect("prepare");
        assert_eq!(
            prepared.source_version().expect("stable source"),
            prepared.prepared_source_version()
        );
    }
}
