//! Private source identity retained by source-backed semantic snapshots.

use litchi_core::SourceVersion;
use litchi_opc::{SourceBackedPackage, SourceLineage};

use crate::error::Result;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SourceProvenance {
    Matched,
    Mismatched,
    Unavailable,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub(crate) struct SourceBinding {
    lineage: Option<SourceLineage>,
    version: Option<SourceVersion>,
}

impl SourceBinding {
    pub(crate) fn capture(package: &SourceBackedPackage) -> Result<Self> {
        package.check_execution()?;
        let version = package.source_version()?;
        package.check_execution()?;
        Ok(Self {
            lineage: Some(package.source_lineage()),
            version: Some(version),
        })
    }

    pub(crate) fn check(&self, package: &SourceBackedPackage) -> Result<SourceProvenance> {
        package.check_execution()?;
        let version = package.source_version()?;
        package.check_execution()?;
        let lineage = package.source_lineage();
        Ok(match (&self.lineage, self.version) {
            (Some(expected_lineage), Some(expected_version)) => {
                if expected_lineage == &lineage && expected_version == version {
                    SourceProvenance::Matched
                } else {
                    SourceProvenance::Mismatched
                }
            },
            _ => SourceProvenance::Unavailable,
        })
    }

    pub(crate) fn same_or_unavailable(&self, other: &Self) -> bool {
        match (&self.lineage, &other.lineage) {
            (Some(left), Some(right)) if left != right => return false,
            _ => {},
        }
        match (self.version, other.version) {
            (Some(left), Some(right)) => left == right,
            _ => true,
        }
    }
}

#[cfg(test)]
mod tests {
    use std::io;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, Ordering};

    use litchi_core::{ReadAt, SourceVersion};
    use litchi_opc::{OpcPackage, PackageWriter};

    use super::*;

    struct VersionedSource {
        bytes: Vec<u8>,
        revision: Arc<AtomicU64>,
    }

    impl ReadAt for VersionedSource {
        fn len(&self) -> io::Result<u64> {
            u64::try_from(self.bytes.len())
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "source too large"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let end = offset.saturating_add(output.len()).min(self.bytes.len());
            output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
            Ok(end - offset)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(17, self.revision.load(Ordering::SeqCst)))
        }
    }

    fn empty_source(revision: Arc<AtomicU64>) -> SourceBackedPackage {
        let bytes = PackageWriter::to_bytes(&OpcPackage::new()).unwrap();
        SourceBackedPackage::from_read_at(Arc::new(VersionedSource { bytes, revision })).unwrap()
    }

    #[test]
    fn binding_distinguishes_matched_foreign_stale_and_unavailable_sources() {
        let revision = Arc::new(AtomicU64::new(0));
        let package = empty_source(Arc::clone(&revision));
        let binding = SourceBinding::capture(&package).unwrap();
        assert_eq!(binding.check(&package).unwrap(), SourceProvenance::Matched);

        let foreign = empty_source(Arc::new(AtomicU64::new(0)));
        assert_eq!(
            binding.check(&foreign).unwrap(),
            SourceProvenance::Mismatched
        );

        let unavailable = SourceBinding::default();
        assert_eq!(
            unavailable.check(&foreign).unwrap(),
            SourceProvenance::Unavailable
        );
        assert!(unavailable.same_or_unavailable(&binding));

        revision.fetch_add(1, Ordering::SeqCst);
        assert!(binding.check(&package).is_err());
    }
}
