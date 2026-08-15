//! Shared ownership for source-backed OPC payloads.

use std::sync::Arc;

use litchi_opc::{OpcError, PartData, SourceBackedPackage};

use crate::error::{Error, Result, allocation};

/// Exact Part bytes without detaching a managed OPC reservation.
#[derive(Clone, Debug)]
pub(crate) enum SourcePayload {
    Managed(PartData),
    Owned(Arc<Vec<u8>>),
}

impl SourcePayload {
    pub(crate) fn from_part_data(package: &SourceBackedPackage, data: PartData) -> Result<Self> {
        if package.cache_diagnostics().budget_managed {
            return Ok(Self::Managed(data));
        }
        Ok(Self::Owned(data.into_arc().map_err(Error::Package)?))
    }

    pub(crate) fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Managed(data) => data.as_bytes(),
            Self::Owned(bytes) => bytes.as_slice(),
        }
    }

    pub(crate) fn len(&self) -> usize {
        self.as_bytes().len()
    }

    /// Return an existing detached allocation without silently escaping a
    /// managed reservation.
    pub(crate) fn detached_arc(&self) -> Result<Arc<Vec<u8>>> {
        match self {
            Self::Managed(_) => Err(Error::Package(OpcError::ManagedPartDataArcEscape)),
            Self::Owned(bytes) => Ok(Arc::clone(bytes)),
        }
    }

    /// Explicitly copy a managed payload under a caller-supplied aggregate
    /// bound. Compatibility payloads keep sharing their allocation.
    pub(crate) fn materialized_arc(
        &self,
        maximum: usize,
        resource: &'static str,
    ) -> Result<Arc<Vec<u8>>> {
        if self.len() > maximum {
            return Err(Error::Invalid(format!(
                "{resource} exceeds the explicit materialization bound {maximum} bytes"
            )));
        }
        match self {
            Self::Managed(data) => copy_bytes(data.as_bytes(), resource),
            Self::Owned(bytes) => Ok(Arc::clone(bytes)),
        }
    }
}

impl PartialEq for SourcePayload {
    fn eq(&self, other: &Self) -> bool {
        self.as_bytes() == other.as_bytes()
    }
}

impl Eq for SourcePayload {}

fn copy_bytes(bytes: &[u8], resource: &'static str) -> Result<Arc<Vec<u8>>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|source| allocation(resource, source))?;
    copy.extend_from_slice(bytes);
    Ok(Arc::new(copy))
}
