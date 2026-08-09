//! Internal codec facade for the web-extension owner.
//!
//! The parent `web` module exposes the semantic model and OPC verbs. This
//! facade keeps the parent-facing symbols stable while the implementation is
//! divided by responsibility.

mod package;
mod relationship;
mod semantic;
mod xml;

#[cfg(test)]
mod tests;

#[cfg(test)]
pub(in crate::web) use semantic::{parse_add_in, parse_panes, write_add_in, write_panes};

#[allow(
    unused_imports,
    reason = "the codec facade preserves validators shared by package and semantic owners"
)]
pub(in crate::web) use super::validation::{
    add_escaped_xml_budget, add_reference_budget, charge_authored_metadata, validate_add_in_budget,
    validate_binding, validate_extension_list, validate_external_uri_reference,
    validate_image_content_type, validate_model, validate_model_with, validate_panes,
    validate_panes_budget, validate_snapshot_resources_with, validate_store_reference,
    validate_task_pane, validate_task_pane_with,
};
#[allow(
    unused_imports,
    reason = "The private codec facade preserves the parent-facing package helper paths."
)]
pub(in crate::web) use package::{
    checked_internal_target, load_snapshot_resources, require_content_type,
};
#[allow(
    unused_imports,
    reason = "The private codec facade preserves the parent-facing relationship helper path."
)]
pub(in crate::web) use relationship::relationship_attr;
#[allow(
    unused_imports,
    reason = "The private codec facade preserves all helpers used by sibling web owners."
)]
pub(in crate::web) use semantic::{
    ParsedPane, enforce_count_with, escape_attr, format_f64, invalid, limit, parse_add_in_with,
    parse_add_in_with_budget, parse_binding, parse_panes_with, parse_panes_with_budget,
    parse_property, parse_store_reference, parse_task_pane, require_nonempty, write_add_in_with,
    write_panes_with, write_store_reference,
};
#[allow(
    unused_imports,
    reason = "The private codec facade preserves all XML helpers used by sibling web owners."
)]
pub(super) use xml::{
    Attribute, ElementEvent, NamespaceScope, Node, NodeFrame, RawFragment, XmlBuildState,
    XmlDocument, attach_node, attr, canonical_node_xml, effective_namespaces, element_children,
    ensure_consumed, escaped_attr_bytes, extension_text_is_allowed, is_drawingml_namespace,
    is_next, next_required, optional_bool_attr, parse_mce_xml, parse_xml, parse_xml_owned,
    push_element, reject_unknown_attributes, require_name, required_attr, retained_namespace_bytes,
    should_capture_extension_list, split_qname,
};
