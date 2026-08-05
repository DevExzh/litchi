//! Temporary host boundary for the canonical PPTX revision owner.

use crate::error::{OoxmlError, Result};
use litchi_opc::OpcPackage;

pub use litchi_pptx::presentation_properties::metadata::revision::{
    Client as ClientRevision, Info as RevisionInformation,
    Namespace as RevisionNamespaceDeclaration, Part as RevisionInformationPart,
};

pub const REVISION_INFORMATION_CONTENT_TYPE: &str =
    litchi_pptx::presentation_properties::metadata::revision::CONTENT_TYPE;
pub const REVISION_INFORMATION_RELATIONSHIP_TYPE: &str =
    litchi_pptx::presentation_properties::metadata::revision::RELATIONSHIP_TYPE;

pub fn load_revision_information(package: &OpcPackage) -> Result<Option<RevisionInformationPart>> {
    litchi_pptx::presentation_properties::metadata::revision::load(package)
        .map_err(OoxmlError::from)
}

pub fn store_revision_information(
    package: &mut OpcPackage,
    value: &RevisionInformationPart,
) -> Result<()> {
    litchi_pptx::presentation_properties::metadata::revision::store(package, value)
        .map_err(OoxmlError::from)
}
