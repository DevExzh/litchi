use crate::{Error, Result};
use litchi_opc::{OpcPackage, PackURI};

/// Inert identity of a VBA project attached to a presentation main part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Project {
    pub(crate) source_part_name: PackURI,
    pub(crate) relationship_id: String,
    pub(crate) project_part_name: PackURI,
}

impl Project {
    pub fn source_part_name(&self) -> &PackURI {
        &self.source_part_name
    }

    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    pub fn project_part_name(&self) -> &PackURI {
        &self.project_part_name
    }

    /// Borrow the opaque `vbaProject.bin` bytes without decoding or copying.
    pub fn payload<'a>(&self, package: &'a OpcPackage) -> Result<&'a [u8]> {
        let part = package.get_part(&self.project_part_name)?;
        if part.content_type() != litchi_opc::constants::content_type::OFC_VBA_PROJECT {
            return Err(Error::ContentType {
                expected: litchi_opc::constants::content_type::OFC_VBA_PROJECT.to_string(),
                actual: part.content_type().to_string(),
            });
        }
        Ok(part.blob())
    }

    pub fn payload_size(&self, package: &OpcPackage) -> Result<usize> {
        Ok(self.payload(package)?.len())
    }
}
