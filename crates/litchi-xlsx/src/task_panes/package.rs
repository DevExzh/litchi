//! OPC graph ownership for persisted task panes.

use litchi_ooxml_common::web::{self as common_web, Conformance, Panes};
use litchi_opc::OpcPackage;

use crate::error::{Result, invalid};

/// Load the package-level task-pane graph, if present.
pub fn load(package: &OpcPackage) -> Result<Option<Panes>> {
    common_web::load(package, &common_web::Limits::standard()).map_err(Into::into)
}

/// Detect the relationship namespace used by an existing task-pane graph.
///
/// MS-OWEXML uses the same package relationship type for the graph while the
/// `r:id` attributes in the XML select Transitional or Strict relationship
/// namespaces. A byte-level namespace probe is sufficient here because the
/// common parser validates the complete graph before a transaction is opened.
pub(crate) fn existing_conformance(package: &OpcPackage) -> Result<Conformance> {
    let mut relationships = package
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == common_web::raw::TASK_PANES_RELATIONSHIP);
    let Some(relationship) = relationships.next() else {
        return Ok(Conformance::Transitional);
    };
    if relationships.next().is_some() {
        return Err(invalid("package has multiple task-pane relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("task-pane relationship cannot be external"));
    }
    let part_name = relationship.target_partname()?;
    let part = package.get_part(&part_name)?;
    let strict_namespace = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
    Ok(
        if part
            .blob()
            .windows(strict_namespace.len())
            .any(|window| window == strict_namespace)
        {
            Conformance::Strict
        } else {
            Conformance::Transitional
        },
    )
}
