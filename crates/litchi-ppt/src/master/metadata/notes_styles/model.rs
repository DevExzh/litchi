//! Typed values for the notes-master `txStyles` package.

use crate::package::Result;

/// Maximum serialized OPC package size accepted by this owner.
pub const MAX_PACKAGE_BYTES: usize = 8 * 1024 * 1024;

/// Maximum number of parts accepted in the inert embedded package.
pub const MAX_PARTS: usize = 16;

/// Maximum XML-part size accepted by the bounded authoring path.
pub const MAX_XML_BYTES: usize = 512 * 1024;

const EMPTY_STYLES: &[u8] =
    br#"<p:txStyles xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;

/// A validated embedded OPC package containing the PresentationML
/// `txStyles` part used by a notes master.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Styles {
    package: Package,
}

impl Styles {
    /// Create a canonical empty `txStyles` package.
    pub fn empty() -> Result<Self> {
        Self::from_xml(EMPTY_STYLES)
    }

    /// Build a deterministic package around one validated `txStyles` XML
    /// document.
    pub fn from_xml(xml: impl AsRef<[u8]>) -> Result<Self> {
        super::package::from_xml(xml.as_ref())
    }

    /// Parse and validate an existing embedded package without normalizing its
    /// bytes.
    pub fn from_package(data: impl Into<Vec<u8>>) -> Result<Self> {
        super::package::from_bytes(data.into())
    }

    /// Borrow the validated package descriptor.
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Borrow the exact package bytes represented by this value.
    pub fn bytes(&self) -> &[u8] {
        self.package.bytes()
    }

    pub(crate) fn from_package_model(package: Package) -> Self {
        Self { package }
    }
}

/// Validated package metadata retained alongside the exact embedded bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Package {
    data: Vec<u8>,
    part_count: usize,
    xml_part_name: String,
}

impl Package {
    pub(crate) fn new(data: Vec<u8>, part_count: usize, xml_part_name: String) -> Self {
        Self {
            data,
            part_count,
            xml_part_name,
        }
    }

    /// Borrow the exact serialized OPC package.
    pub fn bytes(&self) -> &[u8] {
        &self.data
    }

    /// Number of parts in the embedded package.
    pub const fn part_count(&self) -> usize {
        self.part_count
    }

    /// Name of the one PresentationML part containing `txStyles`.
    pub fn xml_part_name(&self) -> &str {
        &self.xml_part_name
    }
}
