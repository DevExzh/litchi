//! Inert SpreadsheetML worksheet-control and ActiveX persistence metadata.
//!
//! ActiveX binaries are deliberately opaque. This module never instantiates a
//! control, resolves a CLSID, executes a macro, decodes MS-OFORMS/CFB data or
//! pictures, or follows an external relationship.

mod codec;
mod model;
mod package;
mod validation;
pub mod vba;

use crate::error::{Error, Result};
use litchi_ooxml_common::Error as CommonError;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SML_STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const XDR_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const AX: &str = "http://schemas.microsoft.com/office/2006/activeX";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const CONTROL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
const CONTROL_REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/control";
const IMAGE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const IMAGE_REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
const BINARY_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
const WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const DESCRIPTOR_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX+xml";
const BINARY_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX";
const MAX_XML: usize = 16 * 1024 * 1024;
const MAX_OUTPUT_XML: usize = 32 * 1024 * 1024;
const MAX_BINARY: usize = 64 * 1024 * 1024;
const MAX_TOTAL_BINARY: usize = 256 * 1024 * 1024;
const MAX_CONTROLS: usize = 65_535;
const MAX_SHAPE_ID: u32 = 67_098_623;
const MAX_CONTROL_NAME_CHARS: usize = 32;
const MAX_PROPERTIES: usize = 65_536;
const MAX_STRING: usize = 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 400_000;

pub use codec::replace_controls_xml;
pub use model::*;
pub use package::{
    load_from_worksheet, remove_from_worksheet, replace_on_worksheet, store_on_worksheet,
};
pub use vba::Project;

fn invalid(value: impl Into<String>) -> Error {
    Error::Invalid(value.into())
}

fn relerr(value: impl Into<String>) -> Error {
    Error::Common(CommonError::Relationship(value.into()))
}

fn limit(value: impl Into<String>) -> Error {
    Error::Invalid(format!("ActiveX resource limit exceeded: {}", value.into()))
}

fn content_type(expected: impl Into<String>, actual: impl Into<String>) -> Error {
    Error::Common(CommonError::ContentType {
        expected: expected.into(),
        actual: actual.into(),
    })
}

fn xml_error(value: impl std::fmt::Display) -> Error {
    Error::Common(CommonError::Xml(value.to_string()))
}
