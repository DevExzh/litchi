//! MS-OWEXML protocol namespaces, relationship types, and safe defaults.

pub(in crate::web) const WEB_EXTENSION_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/webextensions/webextension/2010/11";
pub(in crate::web) const TASK_PANES_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11";
pub(in crate::web) const TRANSITIONAL_RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(in crate::web) const STRICT_RELATIONSHIPS_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(in crate::web) const DRAWINGML_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(in crate::web) const STRICT_DRAWINGML_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/drawingml/main";

pub(in crate::web) const IMAGE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
pub(in crate::web) const STRICT_IMAGE_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image";

pub(in crate::web) const STANDARD_XML_BYTES: usize = 4 * 1024 * 1024;
pub(in crate::web) const STANDARD_TOTAL_XML_BYTES: usize = 64 * 1024 * 1024;
pub(in crate::web) const STANDARD_DEPTH: usize = 128;
pub(in crate::web) const STANDARD_NODES: usize = 65_536;
pub(in crate::web) const STANDARD_ITEMS: usize = 4096;
pub(in crate::web) const STANDARD_STRING_BYTES: usize = 8 * 1024 * 1024;
pub(in crate::web) const STANDARD_TOTAL_STRING_BYTES: usize = 128 * 1024 * 1024;
pub(in crate::web) const STANDARD_IMAGE_BYTES: usize = 64 * 1024 * 1024;
pub(in crate::web) const STANDARD_TOTAL_IMAGE_BYTES: usize = 256 * 1024 * 1024;
pub(in crate::web) const STANDARD_PACKAGE_PARTS: usize = 65_536;
pub(in crate::web) const STANDARD_PACKAGE_RELATIONSHIPS: usize = 262_144;
pub(in crate::web) const STANDARD_PART_ALLOCATIONS: usize = 8_192;
pub(in crate::web) const STANDARD_PART_DELETIONS: usize = 8_192;
