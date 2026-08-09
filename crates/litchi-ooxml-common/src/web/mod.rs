//! Shared, inert Office web-extension and persisted task-pane metadata.

mod codec;
mod model;
mod package;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

#[allow(
    unused_imports,
    reason = "the private web facade retains protocol constants for sibling owners"
)]
pub(in crate::web) use semantic::{
    DRAWINGML_NAMESPACE, IMAGE_RELATIONSHIP_TYPE, STANDARD_DEPTH, STANDARD_IMAGE_BYTES,
    STANDARD_ITEMS, STANDARD_NODES, STANDARD_PACKAGE_PARTS, STANDARD_PACKAGE_RELATIONSHIPS,
    STANDARD_PART_ALLOCATIONS, STANDARD_PART_DELETIONS, STANDARD_STRING_BYTES,
    STANDARD_TOTAL_IMAGE_BYTES, STANDARD_TOTAL_STRING_BYTES, STANDARD_TOTAL_XML_BYTES,
    STANDARD_XML_BYTES, STRICT_DRAWINGML_NAMESPACE, STRICT_IMAGE_RELATIONSHIP_TYPE,
    STRICT_RELATIONSHIPS_NAMESPACE, TASK_PANES_NAMESPACE, TRANSITIONAL_RELATIONSHIPS_NAMESPACE,
    WEB_EXTENSION_NAMESPACE,
};

use crate::{Error, Result};
#[cfg(test)]
use litchi_opc::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};
use std::sync::Arc;

/// Low-level OPC constants for callers constructing synthetic or specialized graphs.
pub mod raw {
    /// Package relationship to the persisted task-pane part.
    pub const TASK_PANES_RELATIONSHIP: &str =
        "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes";
    /// Relationship from the task-pane part to one Office Add-in part.
    pub const ADD_IN_RELATIONSHIP: &str =
        "http://schemas.microsoft.com/office/2011/relationships/webextension";
    /// Content type of a persisted task-pane part.
    pub const TASK_PANES_CONTENT_TYPE: &str = "application/vnd.ms-office.webextensiontaskpanes+xml";
    /// Content type of an Office Add-in part.
    pub const ADD_IN_CONTENT_TYPE: &str = "application/vnd.ms-office.webextension+xml";
}

use raw::{
    ADD_IN_CONTENT_TYPE, ADD_IN_RELATIONSHIP, TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP,
};

pub use model::*;
pub use package::*;
