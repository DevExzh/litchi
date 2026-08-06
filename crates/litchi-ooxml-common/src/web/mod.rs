//! Shared, inert Office web-extension and persisted task-pane metadata.

mod codec;
mod model;
mod package;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

#[allow(unused_imports)]
pub(in crate::web) use semantic::*;

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
