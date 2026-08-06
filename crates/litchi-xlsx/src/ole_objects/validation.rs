//! Safety checks for typed worksheet OLE edits and package graphs.

use litchi_opc::{OpcPackage, PackURI};

use super::model::OleObjects;
use crate::error::Result;

/// Validate a typed OLE graph without requiring package targets.
pub fn objects(value: &OleObjects) -> Result<()> {
    super::model::validate_value(value, false)
}

/// Validate one worksheet's OLE relationships, payloads, and orphan rules.
pub fn graph(package: &OpcPackage, worksheet: &PackURI) -> Result<Option<OleObjects>> {
    super::package::validate_graph(package, worksheet)
}
