//! Host error-boundary adapter for standalone PPTX VBA-project semantics.
//!
//! The owner validates and mutates the PresentationML VBA relationship graph.
//! The small wrapper below preserves the host's payload-decoding helpers for
//! callers that still need an `litchi_vba` project value.

use crate::error::{OoxmlError, Result};
use crate::pptx::media_parts::map_pptx_error;
use litchi_ooxml_common::vba::read_project_part;
use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_vba::{Limits, Payload, project::Project};

/// Relationship metadata for the VBA project attached to a presentation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProject {
    inner: litchi_pptx::presentation::embedded::vba::Project,
}

impl VbaProject {
    /// Return the Presentation part that owns the VBA-project relationship.
    pub fn source_part_name(&self) -> &PackURI {
        self.inner.source_part_name()
    }

    /// Return the relationship ID from the Presentation part to the VBA part.
    pub fn relationship_id(&self) -> &str {
        self.inner.relationship_id()
    }

    /// Return the absolute OPC part name of the VBA project binary part.
    pub fn project_part_name(&self) -> &PackURI {
        self.inner.project_part_name()
    }

    /// Parse the `vbaProject.bin` payload with default resource limits.
    pub fn project(&self, package: &OpcPackage) -> Result<Project> {
        self.project_with(package, &Limits::default())
    }

    /// Parse the `vbaProject.bin` payload with explicit resource limits.
    pub fn project_with(&self, package: &OpcPackage, limits: &Limits) -> Result<Project> {
        read_project_part(package, self.project_part_name(), limits).map_err(OoxmlError::from)
    }
}

/// Discover one structurally conforming Presentation VBA-project relationship.
pub(crate) fn discover_vba_project(
    package: &OpcPackage,
    source: &dyn Part,
) -> Result<Option<VbaProject>> {
    litchi_pptx::presentation::embedded::vba::discover(package, source.partname())
        .map_err(map_pptx_error)
        .map(|project| project.map(|inner| VbaProject { inner }))
}

/// Store an opaque VBA payload through the standalone transactional graph service.
pub(crate) fn store_vba_project(
    package: &mut OpcPackage,
    source: &PackURI,
    payload: Payload,
) -> Result<VbaProject> {
    litchi_pptx::presentation::embedded::vba::store(package, source, payload.into_bytes())
        .map_err(map_pptx_error)
        .map(|inner| VbaProject { inner })
}

/// Remove the complete VBA relationship graph and restore the non-macro main type.
pub(crate) fn remove_vba_project(package: &mut OpcPackage, source: &PackURI) -> Result<bool> {
    litchi_pptx::presentation::embedded::vba::remove(package, source).map_err(map_pptx_error)
}
