//! Temporary host boundary for the canonical PPTX Changes Information owner.

use crate::error::{OoxmlError, Result};
use litchi_opc::OpcPackage;

pub use litchi_pptx::presentation_properties::metadata::changes::{
    Data as ChangesData, Descriptor as ChangeDescriptor, Info as ChangesInformation,
    Kind as ChangeKind, List as ChangesList, Namespace as ChangesNamespaceDeclaration,
    Part as ChangesInformationPart,
};

pub const CHANGES_INFORMATION_CONTENT_TYPE: &str =
    litchi_pptx::presentation_properties::metadata::changes::CONTENT_TYPE;
pub const CHANGES_INFORMATION_RELATIONSHIP_TYPE: &str =
    litchi_pptx::presentation_properties::metadata::changes::RELATIONSHIP_TYPE;

pub fn load_changes_information(package: &OpcPackage) -> Result<Option<ChangesInformationPart>> {
    litchi_pptx::presentation_properties::metadata::changes::load(package).map_err(OoxmlError::from)
}

pub fn store_changes_information(
    package: &mut OpcPackage,
    value: &ChangesInformationPart,
) -> Result<()> {
    litchi_pptx::presentation_properties::metadata::changes::store(package, value)
        .map_err(OoxmlError::from)
}
