//! Inert ActiveX/OCX control metadata attached to slides.

mod codec;
mod model;
mod package;

pub use model::{Binary, Control, Descriptor, Persistence};
pub use package::{Limits, load_slide};

pub(crate) const DESCRIPTOR_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX+xml";
pub(crate) const BINARY_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX";
pub(crate) const ACTIVEX_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/2006/activeX";
pub(crate) const CONTROL_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
pub(crate) const STRICT_CONTROL_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/control";
pub(crate) const BINARY_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
pub(crate) const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_TOTAL_SLIDE_XML_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const MAX_CONTROLS: usize = 4_096;
pub(crate) const MAX_BINARY_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_TOTAL_BINARY_BYTES: usize = 256 * 1024 * 1024;
